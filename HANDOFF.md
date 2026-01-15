# Draft/Deferred Action Feature - Implementation Handoff

## Goal
Allow users to schedule Create/Update actions on draft emails. Actions execute ONLY after email is successfully sent, using the final email body (not draft body). This prevents premature LLM calls and DB writes while email is being edited.

---

## Current Status: PARTIALLY WORKING

### ✅ What Works
1. **Popup compose detection** - Opening Reply/Forward in popup window triggers compose mode
2. **Deferred execution** - When email with UserProperty is sent, SentItems_ItemAdd executes the scheduled action
3. **Data persistence** - UserProperties stores DeferredActionData as JSON on draft emails
4. **LLM + DB execution** - ExecuteDeferredCreateAsync and ExecuteDeferredUpdateAsync work correctly

### ❌ What Doesn't Work
1. **Inline editing NOT detected** - Reply All inline doesn't trigger compose mode (CRITICAL BUG)
2. **Stuck in compose mode** - After closing popup, clicking regular emails doesn't exit compose mode (CRITICAL BUG)
3. **Wrong UI implementation** - Added separate radio buttons instead of making existing buttons latch/toggle (WRONG APPROACH)

---

## Requirements (CORRECT UNDERSTANDING)

### UI Behavior
**In compose mode, existing buttons should become toggle/latching buttons:**
- **Create New** button: Click once = schedule create (button stays "pressed/highlighted"), Click again = cancel schedule
- **Update** button: Click once = schedule update (button stays "pressed/highlighted"), Click again = cancel schedule
- **All other buttons** remain visible and unchanged (Create Multiple, Close, Reopen)
- **NO separate radio button section** - reuse the same button grid

**Compose mode trigger:**
- **Popup Inspector windows** (Reply, Forward, New Email in popup) → Detected via `Inspectors.NewInspector`
- **Inline editing** (Reply All inline, Reply inline) → MUST detect via `Explorer.SelectionChange` checking `mail.Sent == false`

**Compose mode exit:**
- **Popup closed** → Detected via `Inspector.Close` event
- **User clicks regular email** → MUST detect in `OnEmailSelected` and exit compose mode

---

## Technical Details

### Files Modified
1. `Models/DeferredActionData.cs` - Model for storing scheduled action data
2. `ThisAddIn.cs` - Inspector monitoring + SentItems execution
3. `ProjectActionPane.xaml.cs` - Compose mode detection + toggle logic
4. `ProjectActionPane.xaml` - **INCORRECTLY added radio button section (needs removal)**

### Key Code Locations

#### Compose Mode Detection (ThisAddIn.cs)
- **Line 273-317**: `Inspectors_NewInspector` - Detects popup compose windows ✅ WORKS
- **Line 220-271**: `Explorer_SelectionChange` - ATTEMPTED to detect inline compose via `mail.Sent == false` ❌ DOESN'T WORK

#### Compose Mode State (ProjectActionPane.xaml.cs)
- **Line 48-53**: Compose mode fields (`_isComposeMode`, `_composeMail`, `_composeInspector`, `_currentDeferredData`)
- **Line 143-197**: `OnEmailSelected` - ATTEMPTED to exit compose mode ❌ DOESN'T WORK
- **Line 1521-1556**: `OnComposeItemActivated` - Enters compose mode, handles null Inspector ⚠️ PARTIAL
- **Line 1561-1591**: `OnComposeItemDeactivated` - Exits compose mode ⚠️ PARTIAL

#### Deferred Execution (ThisAddIn.cs)
- **Line 300-331**: `SentItems_ItemAdd` - Checks UserProperty and executes ✅ WORKS
- **Line 585-670**: Deferred execution helpers (LoadDeferredData, ClearDeferredData, ExecuteDeferredActionAsync) ✅ WORKS

#### Deferred Action Execution (ProjectActionPane.xaml.cs)
- **Line 1951-2014**: `ExecuteDeferredCreateAsync` - Creates action from sent email ✅ WORKS
- **Line 2020-2081**: `ExecuteDeferredUpdateAsync` - Updates action from sent email ✅ WORKS

---

## Attempted Fixes That FAILED

### Attempt 1: Detect Inline Compose via Explorer_SelectionChange
**Code:** ThisAddIn.cs Line 220-271

```csharp
if (selection is Outlook.MailItem mail)
{
    if (!mail.Sent)  // Check if draft
    {
        _actionPane.OnComposeItemActivated(mail, null); // Enter compose mode
    }
    else
    {
        _actionPane.OnEmailSelected(mail); // Normal mode
    }
}
```

**Why it failed:**
- Inline editing doesn't trigger `Explorer_SelectionChange` event
- When you click Reply All inline, NO selection change event fires
- The mail item remains the same (original received email), not the draft

### Attempt 2: Exit Compose Mode in OnEmailSelected
**Code:** ProjectActionPane.xaml.cs Line 143-197

```csharp
public async void OnEmailSelected(Outlook.MailItem mail)
{
    if (_isComposeMode)
    {
        OnComposeItemDeactivated(); // Exit compose mode
    }
    // ... rest
}
```

**Why it failed:**
- `OnEmailSelected` is never called when popup compose is active
- After closing popup, `OnEmailSelected` doesn't fire automatically
- Need different mechanism to detect mode switching

### Attempt 3: Handle Null Inspector in OnComposeItemActivated
**Code:** ProjectActionPane.xaml.cs Line 1521-1556

```csharp
if (inspector != null)
{
    ((Outlook.InspectorEvents_10_Event)inspector).Close += Inspector_Close;
}
else
{
    System.Diagnostics.Debug.WriteLine("Inline compose - no Inspector");
}
```

**Why it's partial:**
- Works for popup windows (inspector != null) ✅
- But inline compose is never detected in the first place ❌
- So this code path for null Inspector never executes

---

## Incorrect UI Implementation

### What Was Built (WRONG)
Added new "Schedule Action" section with 3 radio buttons:
- `ProjectActionPane.xaml` Line 240-283: Radio button section
- `ProjectActionPane.xaml.cs` Line 1429-1487: `ScheduleRadio_Changed` handler

**This is WRONG because:**
- User wanted existing buttons to become toggle/latch buttons
- Separate radio buttons add UI clutter
- Doesn't match the natural workflow

### What Should Be Built (CORRECT)
Existing buttons should change behavior in compose mode:
- **Normal mode**: Click Create → Execute immediately
- **Compose mode**: Click Create → Toggle scheduled state (button highlights/un-highlights)
- Button visual states:
  - Unscheduled: Normal blue button (existing style)
  - Scheduled: Green/highlighted button (different background)
  - Text changes: "Create New" → "Cancel Create" when scheduled

---

## Root Cause Analysis

### Why Inline Editing Not Detected

**Problem:** Outlook inline editing doesn't create an Inspector window

**What Outlook does:**
1. User clicks "Reply All" inline
2. Reading pane shows inline compose box
3. **NO** new Inspector window created
4. **NO** `Inspectors.NewInspector` event fires
5. **NO** `Explorer_SelectionChange` event fires (selection remains the original email)

**Possible solutions to try:**
1. Hook `MailItem.Reply` / `MailItem.ReplyAll` / `MailItem.Forward` events on selected mail
2. Monitor `ActiveExplorer.ActiveInlineResponse` property (if exists)
3. Poll for inline response state in a timer
4. Use UIAutomation to detect inline compose UI elements (HACKY)
5. Hook `Application.ItemSend` event and backtrack to figure out compose mode (TOO LATE)

### Why Stuck in Compose Mode

**Problem:** No event fires when user switches from compose to regular email

**What happens:**
1. User opens compose popup → `_isComposeMode = true`
2. User closes popup → `Inspector.Close` fires → `_isComposeMode = false` ✅
3. BUT if user **doesn't close popup** and just clicks another email:
   - `Explorer_SelectionChange` doesn't fire (popup has focus, not Explorer)
   - `OnEmailSelected` doesn't fire
   - `_isComposeMode` stays `true` ❌

**Possible solutions to try:**
1. Hook `Application.Explorers` and monitor `ActiveExplorer` changes
2. Hook `Inspector.Deactivate` event (when popup loses focus)
3. In `Explorer_SelectionChange`, always clear compose mode first before processing
4. Track active window and detect when Explorer regains focus

---

## Correct Implementation Steps

### Step 1: Fix Inline Compose Detection
Try hooking MailItem events:
```csharp
// In Explorer_SelectionChange, after getting mail:
if (selection is Outlook.MailItem mail)
{
    // Hook compose events
    mail.Reply += MailItem_Reply;
    mail.ReplyAll += MailItem_ReplyAll;
    mail.Forward += MailItem_Forward;
}

// Event handlers
private void MailItem_Reply(object response, ref bool cancel)
{
    if (response is Outlook.MailItem draft)
    {
        _actionPane.OnComposeItemActivated(draft, null);
    }
}
```

### Step 2: Fix Compose Mode Exit
Always clear compose mode in `Explorer_SelectionChange`:
```csharp
private void Explorer_SelectionChange()
{
    // ALWAYS clear compose mode when Explorer selection changes
    if (_isComposeMode)
    {
        _actionPane.Dispatcher.Invoke(() =>
        {
            _actionPane.OnComposeItemDeactivated();
        });
    }

    // Then process new selection...
}
```

### Step 3: Fix Button UI (Remove Radio Buttons)
1. **Remove:** `ScheduleActionSection` Border from XAML
2. **Remove:** `ScheduleRadio_Changed` handler
3. **Restore:** Original button click handlers with toggle logic:

```csharp
private void CreateButton_Click(object sender, RoutedEventArgs e)
{
    if (_isComposeMode)
    {
        // TOGGLE scheduled create
        if (_currentDeferredData?.Mode == "Create")
        {
            // Cancel
            ClearDeferredData(_composeMail);
            _currentDeferredData = null;
            CreateButton.Background = PrimaryAccent; // Blue
            CreateButton.Content = "Create New";
        }
        else
        {
            // Schedule
            var data = new DeferredActionData { Mode = "Create" };
            SaveDeferredData(_composeMail, data);
            _currentDeferredData = data;
            CreateButton.Background = Green; // Highlight
            CreateButton.Content = "Cancel Create";
        }
    }
    else
    {
        // Normal immediate execution
        ProcessCreateActionAsync(_currentMail);
    }
}
```

### Step 4: Visual Feedback
Update `UpdateUIForComposeMode` to set button states:
```csharp
private void UpdateUIForComposeMode()
{
    if (_isComposeMode)
    {
        // Update button visuals based on scheduled state
        if (_currentDeferredData?.Mode == "Create")
        {
            CreateButton.Background = Green;
            CreateButton.Content = "Cancel Create";
        }
        else
        {
            CreateButton.Background = PrimaryAccent;
            CreateButton.Content = "Create New";
        }

        if (_currentDeferredData?.Mode == "Update")
        {
            UpdateButton.Background = Green;
            UpdateButton.Content = "Cancel Update";
        }
        else
        {
            UpdateButton.Background = PrimaryAccent;
            UpdateButton.Content = "Update";
        }
    }
    else
    {
        // Restore normal state
        CreateButton.Background = PrimaryAccent;
        CreateButton.Content = "Create New";
        UpdateButton.Background = PrimaryAccent;
        UpdateButton.Content = "Update";
    }
}
```

---

## Testing Checklist

### Priority 1: Critical Bugs
- [ ] Click Reply All inline → Compose mode MUST activate
- [ ] Click Reply inline → Compose mode MUST activate
- [ ] Open compose popup, close it, click regular email → Compose mode MUST deactivate

### Priority 2: Button Toggle Behavior
- [ ] In compose mode: Click Create → Button highlights green, text = "Cancel Create"
- [ ] Click Create again → Button returns to blue, text = "Create New"
- [ ] In compose mode: Click Update (with action selected) → Button highlights green, text = "Cancel Update"
- [ ] Click Update again → Button returns to blue, text = "Update"

### Priority 3: Deferred Execution
- [ ] Schedule create, send email → Action created in Google Sheets
- [ ] Schedule update, send email → Action updated in Google Sheets
- [ ] Don't schedule, send email → Nothing happens (normal send)

---

## Known Issues

1. **Inline compose detection is fundamentally difficult** in Outlook VSTO
   - No reliable event fires for inline Reply/Forward
   - May need to poll or use workarounds

2. **Compose mode state management is fragile**
   - Multiple Inspector windows can exist simultaneously
   - Need better tracking of which Inspector/draft is active

3. **UI complexity with toggle buttons**
   - Need to track scheduled state per button
   - Mutually exclusive scheduling (Create OR Update, not both)

---

## Recommendations for Next Developer

1. **Research Outlook inline compose detection**
   - Search for "Outlook VSTO detect inline reply"
   - Check if `ActiveInlineResponse` or `ActiveInlineResponseItem` exists
   - Consider polling approach as last resort

2. **Simplify compose mode tracking**
   - Instead of tracking `_composeInspector`, track list of active compose MailItems
   - Use MailItem GUID or EntryID as key

3. **Test extensively with Debug output**
   - Add Debug.WriteLine to EVERY event handler
   - Monitor which events fire in different scenarios
   - Document findings in this file

4. **Consider alternative approaches**
   - Instead of detecting compose mode, detect at Send time
   - Show scheduling UI in a separate dialog when user clicks Create/Update
   - Use Outlook Form Regions for draft-specific UI

---

## References

- **Plan file:** `C:\Users\wally\.claude\plans\lucky-plotting-bengio.md`
- **Original requirements:** `FastPMHelperAddin\Prompt.txt`
- **Agent exploration report:** See conversation history for detailed code analysis (agentId: abfdfdb)

---

## Contact

If continuing this work, start by:
1. Reading this handoff document completely
2. Testing current behavior with Debug output enabled
3. Researching Outlook inline compose detection methods
4. Implementing one fix at a time and testing thoroughly

Good luck! 🚀
