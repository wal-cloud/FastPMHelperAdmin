# HANDOFF DOCUMENT - Outbox Stuck Email Bug Investigation

**Date**: 2026-01-21
**Project**: FastPMHelperAddin (VSTO Outlook Add-in)
**Issue**: Emails get stuck in Outbox when 2-minute delay rule is enabled
**Status**: INVESTIGATING - Root cause identified but not fully resolved

---

## 1. PROBLEM DESCRIPTION

### User-Reported Behavior
When the user has an Outlook rule that delays all outbound emails by 2 minutes:
- Emails get stuck in Outbox and never send
- This happens **even when the user does NOT tag the email** with any deferred action (Create/Update/Close/Reopen)
- The stuck email shows:
  - **Date: NONE** (instead of today's date)
  - **Not in italics** (Outlook treats it as "open" or being edited)
- Normal emails that will send correctly show today's date and are in italics

### Impact
- Feature is unusable for users with delay rules
- Could also affect users with no internet (emails temporarily in Outbox)
- The add-in appears to be interfering with Outlook's delayed send mechanism

---

## 2. ROOT CAUSE ANALYSIS

### The Core Issue
The add-in accesses `mail.UserProperties` on emails that Outlook reports as being in the "Outbox" folder. This access appears to interfere with Outlook's ability to process the delayed send.

### Key Discovery: Folder Detection Anomaly
When composing an **inline reply**, the draft's `mail.Parent` property returns **"Outbox"** as the folder name, even though:
- `mail.Submitted = False` (not yet submitted for sending)
- `mail.Sent = False` (not yet sent)
- The user is still actively composing the email

This is unexpected - inline drafts should typically be in "Drafts" or have no folder.

### Debug Evidence
From `Error.txt` logs:
```
=== Explorer_InlineResponse EVENT FIRED === (Time: 09:31:55.496)
  Inline compose started for: RE: B&R25054 | AMM25012-PPN-311 |  TQ010 on bellows
  Submitted status: False
  Entering compose mode (inline)
OnComposeItemActivated: RE: B&R25054 | AMM25012-PPN-311 |  TQ010 on bellows
  Inspector: Inline editing
[OUTBOX-DEBUG] Sidebar LoadDeferredData called:
[OUTBOX-DEBUG]   Subject: RE: B&R25054 | AMM25012-PPN-311 |  TQ010 on bellows
[OUTBOX-DEBUG]   Folder: Outbox                          <-- ANOMALY: Draft shows as Outbox
[OUTBOX-DEBUG]   State: Sent=False, Submitted=False
[OUTBOX-DEBUG] WARNING: SIDEBAR ACCESSING OUTBOX ITEM!
```

After sending, the email shows in Outbox with `Submitted=True`:
```
=== Explorer_SelectionChange EVENT FIRED === (Time: 09:32:10.761)
  Email selected: RE: B&R25054 | AMM25012-PPN-311 |  TQ010 on bellows
  Sent status: False
  Submitted status: True
  [OUTBOX-DEBUG] Folder: Outbox
  [OUTBOX-DEBUG] OUTBOX ITEM SELECTED!
```

### Theory
1. When composing inline, the draft is somehow associated with "Outbox" folder
2. The add-in calls `LoadDeferredData()` which accesses `mail.UserProperties`
3. This access (even read-only) may "touch" the email in a way that:
   - Marks it as being edited
   - Prevents Outlook's delay rule from processing it
   - Causes the "NONE" date and non-italic display

---

## 3. WHAT WE TRIED

### Attempt 1: Skip items with `Submitted=True`
**Location**: `ThisAddIn.cs:Explorer_SelectionChange()` and `Explorer_InlineResponse()`

**Code Added**:
```csharp
// Skip if Submitted (truly in Outbox awaiting send)
if (mail.Submitted)
{
    System.Diagnostics.Debug.WriteLine($"  [OUTBOX-DEBUG] SKIPPING - Submitted=True");
    return;
}
```

**Result**: PARTIAL - This correctly skips emails that are already in Outbox awaiting send, but doesn't prevent the initial access during compose.

### Attempt 2: Skip items in Outbox folder (REVERTED)
**Location**: `ThisAddIn.cs`, `ProjectActionPane.xaml.cs`

**Code Added**:
```csharp
var folder = mail.Parent as Outlook.MAPIFolder;
if (folder?.Name?.Equals("Outbox", StringComparison.OrdinalIgnoreCase) == true)
{
    return; // Skip Outbox items
}
```

**Result**: FAILED - This broke the deferred action feature because:
- Inline drafts report "Outbox" as their folder even when `Submitted=False`
- Skipping these items prevents users from using the tagging buttons during inline compose
- **REVERTED** - This approach is too aggressive

### Attempt 3: Add comprehensive debug logging
**Locations**:
- `InspectorComposeRibbon.cs:LoadDeferredData()` - lines 995-1058
- `InspectorComposeRibbon.cs:SaveDeferredData()` - lines 955-990
- `InspectorComposeRibbon.cs:ClearDeferredData()` - lines 1063-1105
- `ProjectActionPane.xaml.cs:LoadDeferredData()` - lines 2370-2420
- `ThisAddIn.cs:Explorer_SelectionChange()` - lines 246-270
- `ThisAddIn.cs:Explorer_InlineResponse()` - lines 322-338

**Debug Output Format**:
```
[OUTBOX-DEBUG] LoadDeferredData called:
[OUTBOX-DEBUG]   Subject: {subject}
[OUTBOX-DEBUG]   Folder: {folderName}
[OUTBOX-DEBUG]   State: Sent={sent}, Submitted={submitted}
```

**Result**: Successfully identified the folder detection anomaly (see Section 2)

---

## 4. CURRENT STATE OF CODE

### Debug Logging Added (Still Active)
The following debug logging is still in place for future investigation:

**InspectorComposeRibbon.cs:LoadDeferredData()** (~line 1000):
```csharp
// DEBUG: Log folder and mail state to detect Outbox access
string folderName = "(unknown)";
string mailState = "(unknown)";
try
{
    var folder = mail.Parent as Outlook.MAPIFolder;
    folderName = folder?.Name ?? "(no folder)";
    mailState = $"Sent={mail.Sent}, Submitted={mail.Submitted}";
}
catch { /* ignore folder access errors */ }

System.Diagnostics.Debug.WriteLine($"[OUTBOX-DEBUG] LoadDeferredData called:");
System.Diagnostics.Debug.WriteLine($"[OUTBOX-DEBUG]   Subject: {mail.Subject}");
System.Diagnostics.Debug.WriteLine($"[OUTBOX-DEBUG]   Folder: {folderName}");
System.Diagnostics.Debug.WriteLine($"[OUTBOX-DEBUG]   State: {mailState}");
System.Diagnostics.Debug.WriteLine($"[OUTBOX-DEBUG]   StackTrace: {Environment.StackTrace}");

// WARN if accessing Outbox item
if (folderName.Equals("Outbox", StringComparison.OrdinalIgnoreCase))
{
    System.Diagnostics.Debug.WriteLine($"[OUTBOX-DEBUG] WARNING: ACCESSING OUTBOX ITEM!");
}
```

**Similar logging in**:
- `SaveDeferredData()` - warns on Outbox writes
- `ClearDeferredData()` - warns on Outbox clears
- `ProjectActionPane.xaml.cs:LoadDeferredData()` - sidebar version

### Skip Logic Added (Still Active)
**ThisAddIn.cs:Explorer_SelectionChange()** (~line 265):
```csharp
// Skip if Submitted (truly in Outbox awaiting send)
if (mail.Submitted)
{
    System.Diagnostics.Debug.WriteLine($"  [OUTBOX-DEBUG] SKIPPING - Submitted=True");
    return;
}
```

**ThisAddIn.cs:Explorer_InlineResponse()** (~line 334):
```csharp
// Skip if already submitted (in Outbox)
if (draft.Submitted)
{
    System.Diagnostics.Debug.WriteLine($"  [OUTBOX-DEBUG] Skipping - email already submitted");
    return;
}
```

---

## 5. THEORIES FOR ROOT CAUSE

### Theory 1: UserProperties Access Locks the Email
Accessing `mail.UserProperties` (even read-only via `.Find()`) may:
- Open the email for editing internally
- Set a flag that prevents Outlook's rules from processing it
- Conflict with the delay rule's internal locking mechanism

**Evidence**: The email shows as non-italic (open/being edited) in Outbox

### Theory 2: mail.Parent Access Causes Issues
Just reading `mail.Parent` to get the folder might be enough to interfere with the email.

**Test Needed**: Remove all `mail.Parent` access and see if issue persists

### Theory 3: Timing/Race Condition
The add-in accesses the email during the brief window when:
1. User clicks Send
2. Email moves to Outbox
3. Delay rule timer starts
4. Add-in's `SelectionChange` fires and accesses the email
5. This access disrupts the delay rule

**Evidence**: `SelectionChange` fires multiple times rapidly after sending

### Theory 4: COM Reference Holding
The add-in stores `_composeMail = mail` which holds a COM reference.
Even after `_composeMail = null`, GC may not immediately release it.

**Potential Fix**: Use `Marshal.ReleaseComObject(_composeMail)` explicitly

---

## 6. POTENTIAL SOLUTIONS TO EXPLORE

### Solution 1: Delay UserProperties Access
Don't access `UserProperties` until the user actually clicks a tagging button:
```csharp
// Current (problematic):
public void OnComposeItemActivated(...)
{
    _currentDeferredData = LoadDeferredData(mail); // Immediate access
}

// Proposed:
public void OnComposeItemActivated(...)
{
    _deferredDataLoaded = false; // Lazy load later
}

private DeferredActionData GetDeferredData()
{
    if (!_deferredDataLoaded)
    {
        _currentDeferredData = LoadDeferredData(_composeMail);
        _deferredDataLoaded = true;
    }
    return _currentDeferredData;
}
```

### Solution 2: Use EntryID-based Detection Instead of Folder
Instead of checking `mail.Parent` for folder, use `mail.EntryID` patterns or other properties.

### Solution 3: Check for Delay Rule Specifically
Detect if the email is subject to a delay rule and avoid accessing it:
```csharp
// Check if email has deferred delivery time set
if (mail.DeferredDeliveryTime > DateTime.Now)
{
    // Skip - email is being delayed by a rule
    return;
}
```

### Solution 4: Use Application.ItemSend Event
Instead of detecting compose mode via `SelectionChange`, use:
```csharp
_app.ItemSend += Application_ItemSend;

private void Application_ItemSend(object item, ref bool cancel)
{
    // Process deferred actions here, BEFORE email goes to Outbox
}
```

### Solution 5: Move All Logic to SentItems_ItemAdd
Only process emails AFTER they successfully send:
- Remove all compose-time access
- Only read/write UserProperties in `SentItems_ItemAdd`
- Downside: Loses the ability to show UI state during compose

---

## 7. FILES MODIFIED IN THIS SESSION

### ThisAddIn.cs
- Added `Submitted` status logging
- Added folder detection logging
- Added skip logic for `Submitted=True` items
- Lines affected: ~246-320

### InspectorComposeRibbon.cs
- Added Outbox debug logging to `LoadDeferredData()` (~line 995-1058)
- Added Outbox debug logging to `SaveDeferredData()` (~line 955-990)
- Added Outbox debug logging to `ClearDeferredData()` (~line 1063-1105)

### ProjectActionPane.xaml.cs
- Added Outbox debug logging to `LoadDeferredData()` (~line 2370-2420)

---

## 8. HOW TO TEST

### Test Setup
1. Enable a 2-minute delay rule on all outbound emails in Outlook
2. Build the add-in:
```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "C:\Users\wally\source\repos\FastPMHelperAddin\FastPMHelperAddin.sln" /t:Rebuild /p:Configuration=Debug
```
3. Start Outlook
4. Open `Error.txt` to monitor debug output

### Test Case 1: Inline Reply (No Tagging)
1. Select an email in Inbox
2. Click Reply (inline)
3. Type a response
4. Click Send
5. **Expected**: Email goes to Outbox, waits 2 min, sends
6. **Bug Behavior**: Email stuck in Outbox with Date=NONE, not italic

### Test Case 2: Inline Reply (With Tagging)
1. Select an email in Inbox
2. Click Reply (inline)
3. Click "Create" button in sidebar
4. Click Send
5. **Expected**: Email goes to Outbox, waits 2 min, sends, action created
6. **Bug Behavior**: Email stuck in Outbox

### Test Case 3: Pop-out Reply
1. Select an email in Inbox
2. Click Reply, then Pop Out
3. Type a response
4. Click Send
5. Check if behavior differs from inline

### What to Look For in Debug Output
- `[OUTBOX-DEBUG]` lines showing folder access
- Whether `Folder: Outbox` appears for compose drafts
- Stack traces showing where access originates
- Timing of `SelectionChange` events relative to send

---

## 9. RELATED CONTEXT

### DeferredActionData Model
```csharp
public class DeferredActionData
{
    public string Mode { get; set; }        // "Create", "Update", "Close", "Reopen", "CreateMultiple"
    public int? ActionID { get; set; }      // Linked action ID
    public string ManualTitle { get; set; }
    public string ManualAssignee { get; set; }
}
```

### UserProperty Name
`"FastPMDeferredAction"` - stored in `mail.UserProperties`

### Execution Flow
1. User toggles button (Create/Update/etc.) in compose mode
2. `SaveDeferredData()` writes JSON to `mail.UserProperties["FastPMDeferredAction"]`
3. User sends email
4. Email arrives in Sent Items
5. `SentItems_ItemAdd` event fires
6. `LoadDeferredData()` reads the property
7. `ExecuteDeferredActionAsync()` performs the action
8. `ClearDeferredData()` removes the property

---

## 10. KNOWN WORKING SCENARIOS

- Add-in works correctly when delay rule is **disabled**
- Pop-out compose windows work correctly (ribbon mode)
- Deferred actions execute correctly when email successfully sends
- Sidebar releases correctly when pop-out is detected

---

## 11. BUILD INSTRUCTIONS

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "C:\Users\wally\source\repos\FastPMHelperAddin\FastPMHelperAddin.sln" /t:Rebuild /p:Configuration=Debug
```

**DO NOT USE** `dotnet build` - it fails with MSB4019 error.

---

## 12. DEBUG OUTPUT LOCATION

`C:\Users\wally\source\repos\FastPMHelperAddin\FastPMHelperAddin\Error.txt`

---

**End of Handoff Document**
**Priority**: HIGH - Feature breaks for users with delay rules
**Next Steps**: Implement one of the solutions in Section 6 and test
