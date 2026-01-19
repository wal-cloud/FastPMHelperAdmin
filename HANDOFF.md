# HANDOFF DOCUMENT - Inspector Ribbon State Persistence Bug

**Date**: 2026-01-19
**Project**: FastPMHelperAddin (VSTO Outlook Add-in)
**Feature**: Hybrid Architecture for Compose Window Action Tracking
**Status**: 🔴 BLOCKED - Critical state persistence bug unresolved

---

## 1. FEATURE OVERVIEW

### What We Built
A **Hybrid Architecture** to fix a "sticky sidebar bug" where the Sidebar gets locked onto popped-out Inspector windows and can't track the main Explorer window anymore.

### Architecture Design
- **Main Window (Explorer)**: Sidebar handles Read Mode + Inline Compose
- **Pop-out Windows (Inspectors)**: Custom Ribbon handles Pop-out Compose
- **Critical Rule**: Sidebar NEVER controls pop-out windows - releases immediately when pop-out detected

### Components Implemented
1. **InspectorWrapper.cs** (NEW) - Manages individual Inspector window lifecycle
2. **InspectorComposeRibbon.xml** (NEW) - Ribbon UI definition with action tracking controls
3. **InspectorComposeRibbon.cs** (NEW) - Ribbon code-behind with UserProperty persistence
4. **ThisAddIn.cs** (MODIFIED) - Added wrapper dictionary and sidebar release logic
5. **ProjectActionPane.xaml.cs** (MODIFIED) - Exposed public properties for ribbon access
6. **FastPMHelperAddin.csproj** (MODIFIED) - Added compile and embedded resource entries

### Ribbon Features
- **5 Toggle Buttons**: Create, Create Multiple, Update, Close, Reopen
- **Action Dropdown**: Grouped actions (Linked, Package, Project, Other) with auto-selection
- **Status Label**: Shows current scheduled action
- **UserProperty Persistence**: Stores `DeferredActionData` on mail item using `"FastPMDeferredAction"` property
- **Mutual Exclusivity**: Only one action can be scheduled at a time

---

## 2. WHAT WORKS ✅

### Successfully Implemented
1. ✅ **Sidebar Release on Pop-out** - `ThisAddIn.Inspectors_NewInspector` correctly calls `OnComposeItemDeactivated()` to release sidebar
2. ✅ **Inspector Lifecycle Management** - InspectorWrapper properly tracks and cleans up Inspector windows
3. ✅ **Ribbon Visibility** - Ribbon appears in compose windows using `TabNewMailMessage` (NOT `TabMessage`)
4. ✅ **Action Dropdown** - Successfully populates with grouped actions using callback pattern (`getItemCount`, `getItemLabel`, `getItemID`)
5. ✅ **Toggle Buttons** - All 5 toggles functional, save/load from UserProperty correctly
6. ✅ **Mutual Exclusivity** - Only one action mode can be scheduled at a time
7. ✅ **UserProperty Persistence** - Data saves to mail item and persists across sessions
8. ✅ **Action Execution** - When email is sent, deferred action executes correctly in `SentItems_ItemAdd`
9. ✅ **First Window Behavior** - First compose window works perfectly:
   - Auto-selects first linked action
   - Toggle buttons work
   - Action executes on send

---

## 3. THE CRITICAL BUG 🔴

### Problem Description
**State (both dropdown selection AND toggle button state) persists across different compose windows.**

### User-Reported Behavior
> "Great that worked the first time but then it got stuck on that. So it had that linked action, + I ticked update, hit send, action updated correctly. But then every new draft I created (no matter what the actual linked action was) had that original linked action AND update already preselected"

### Reproduction Steps
1. Open compose window #1 (Reply to Email A)
   - ✅ Dropdown auto-selects first linked action (e.g., "Action 123")
   - ✅ Click "Update" toggle - saves to UserProperty
   - ✅ Send email - action executes correctly
2. Open compose window #2 (Reply to Email B - completely different email)
   - ❌ **BUG**: Dropdown still shows "Action 123" selected
   - ❌ **BUG**: "Update" toggle still pressed
   - 🤔 Expected: Clean state, auto-select first linked action for Email B

### Impact
Every new compose window inherits state from the previous window, making the feature unusable for multiple compose sessions.

---

## 4. DEBUGGING ATTEMPTS & FINDINGS

### Attempt 1: State Reset in Ribbon_Load
**Hypothesis**: Cached state carrying over
**Action**: Added comprehensive state reset in `Ribbon_Load()`:
```csharp
_selectedAction = null;
_currentDeferredData = null;
_dropdownActions.Clear();
_dropdownLabels.Clear();
_linkedActionsCount = 0;
```
**Result**: ❌ FAILED - State still persists
**User Feedback**: "It is still persisting both the action AND the state."

### Attempt 2: Made All GetPressed Callbacks Stateless
**Hypothesis**: Callbacks using stale cached data
**Action**: Changed all `GetPressed` callbacks to load fresh from UserProperty:
```csharp
public bool GetUpdateActionPressed(Office.IRibbonControl control)
{
    var mail = GetCurrentMailItem();
    if (mail == null) return false;
    var data = LoadDeferredData(mail);
    return data?.Mode == "Update";
}
```
**Result**: ❌ FAILED - State still persists
**User Feedback**: "It is still doing the same thing"

### Attempt 3: Made GetSelectedActionIndex Stateless
**Hypothesis**: Dropdown selection using cached `_selectedAction` field
**Action**: Rewrote `GetSelectedActionIndex()` to prioritize fresh UserProperty data:
```csharp
public int GetSelectedActionIndex(Office.IRibbonControl control)
{
    var mail = GetCurrentMailItem();
    if (mail != null)
    {
        var data = LoadDeferredData(mail);
        if (data?.ActionID.HasValue == true)
        {
            // Find and return index from saved data
            for (int i = 0; i < _dropdownActions.Count; i++)
            {
                if (_dropdownActions[i].Id == data.ActionID.Value)
                    return i;
            }
        }
    }
    // Auto-select first linked action if no saved data
    if (_linkedActionsCount > 0) return 0;
    return -1;
}
```
**Result**: ❌ FAILED - State still persists

### Attempt 4: Clear _selectedAction on Toggle Cancel
**Hypothesis**: `_selectedAction` field not being cleared
**Action**: Added `_selectedAction = null;` to all toggle cancel operations
**Result**: ❌ FAILED - State still persists

### Attempt 5: Added Comprehensive Debugging (IN PROGRESS)
**Hypothesis**: Need to understand ribbon instance lifecycle
**Action**: Added extensive logging to track:
- Instance creation with unique IDs (`Ribbon-1`, `Ribbon-2`, etc.)
- Which instance `Ribbon_Load()` is called on
- Which instance callbacks are executed on
- What data each instance sees

**Code Added**:
```csharp
// InspectorComposeRibbon.cs
private static int _instanceCounter = 0;
private readonly int _instanceId;
public string InstanceID => $"Ribbon-{_instanceId}";

public InspectorComposeRibbon()
{
    _instanceId = System.Threading.Interlocked.Increment(ref _instanceCounter);
    System.Diagnostics.Debug.WriteLine($"*** InspectorComposeRibbon Constructor: {InstanceID} ***");
}

// ThisAddIn.cs
protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
{
    System.Diagnostics.Debug.WriteLine($"=== CreateRibbonExtensibilityObject CALLED at {DateTime.Now:HH:mm:ss.fff} ===");
    var ribbon = new InspectorComposeRibbon();
    System.Diagnostics.Debug.WriteLine($"    Created new ribbon instance: {ribbon.InstanceID}");
    return ribbon;
}
```

**Status**: ⏳ NOT YET TESTED - User interrupted before rebuild/test
**Next Step**: Rebuild and reproduce issue to see debug output

---

## 5. THEORIES & HYPOTHESES

### Theory 1: Outlook Reuses Ribbon Instances (MOST LIKELY)
**Evidence**:
- All stateless attempts failed
- State persists despite fresh UserProperty reads
- Callbacks may be executing on wrong ribbon instance

**What to Check**:
- Does `CreateRibbonExtensibilityObject()` get called once or per window?
- Are callbacks executing on the same instance ID across different windows?
- Is Outlook caching ribbon state somewhere?

### Theory 2: UserProperty Data Actually Persists
**Evidence**:
- `LoadDeferredData()` might be reading from a different mail item than expected
- `GetCurrentMailItem()` might be returning wrong reference

**What to Check**:
- Does new compose window have existing `FastPMDeferredAction` property?
- Is `mail.EntryID` different between windows?
- Is `GetCurrentMailItem()` returning the active Inspector's mail or a stale reference?

### Theory 3: Outlook Ribbon Cache
**Evidence**:
- Ribbon XML might be cached with state
- `Invalidate()` might not fully refresh state

**What to Check**:
- Does calling `_ribbon.Invalidate()` actually force fresh callback execution?
- Is there a global ribbon cache we need to clear?

---

## 6. KEY FILES & LOCATIONS

### New Files Created
```
FastPMHelperAddin/
├── InspectorWrapper.cs                  (NEW - Inspector lifecycle manager)
├── InspectorComposeRibbon.cs            (NEW - Ribbon code-behind, ~870 lines)
└── InspectorComposeRibbon.xml           (NEW - Ribbon UI definition)
```

### Modified Files
```
FastPMHelperAddin/
├── ThisAddIn.cs                         (MODIFIED - Lines 23, 346-401, 737-745)
│   ├── Added: _inspectorWrappers dictionary
│   ├── Modified: Inspectors_NewInspector (THE CRITICAL FIX)
│   └── Added: CreateRibbonExtensibilityObject override
├── UI/ProjectActionPane.xaml.cs         (MODIFIED - Lines 34-39)
│   └── Added: Public properties for ribbon access
└── FastPMHelperAddin.csproj             (MODIFIED - Lines ~305, ~381)
    ├── Added: <Compile Include="InspectorWrapper.cs" />
    ├── Added: <Compile Include="InspectorComposeRibbon.cs" />
    └── Added: <EmbeddedResource Include="InspectorComposeRibbon.xml" />
```

### Critical Code Section - ThisAddIn.cs Lines 346-401
This is the **PRIMARY FIX** for the sticky sidebar bug:
```csharp
private void Inspectors_NewInspector(Outlook.Inspector inspector)
{
    if (inspector.CurrentItem is Outlook.MailItem mail && !mail.Sent)
    {
        System.Diagnostics.Debug.WriteLine("  Detected compose window - creating InspectorWrapper");

        // Create InspectorWrapper
        var wrapper = new InspectorWrapper(inspector);
        string key = GetInspectorKey(inspector);
        _inspectorWrappers[key] = wrapper;

        // CRITICAL: Release Sidebar immediately (don't call OnComposeItemActivated!)
        _actionPane.Dispatcher.Invoke(() =>
        {
            if (_actionPane.IsComposeMode)
            {
                System.Diagnostics.Debug.WriteLine("  Sidebar was in compose mode - releasing to Ribbon");
                _actionPane.OnComposeItemDeactivated();
            }
        });
    }
}
```

**Why This Works**: Sidebar immediately releases when Inspector opens, preventing it from getting "stuck" tracking the Inspector.

### Critical Code Section - InspectorComposeRibbon.cs
**UserProperty Management** (Lines 752-831):
- `SaveDeferredData()` - Serializes to JSON, stores in UserProperty
- `LoadDeferredData()` - Reads UserProperty, deserializes JSON
- `ClearDeferredData()` - Deletes UserProperty

**Property Name**: `"FastPMDeferredAction"` (MUST match sidebar)

---

## 7. TECHNICAL DETAILS

### Build System
**Project Type**: VSTO Add-in (.NET Framework) - Old-style .csproj

**CRITICAL**: Must use full MSBuild path, NOT `dotnet build`:
```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "C:\Users\wally\source\repos\FastPMHelperAddin\FastPMHelperAddin.sln" /t:Rebuild /p:Configuration=Debug
```

### Ribbon XML Configuration
**Tab**: `TabNewMailMessage` (NOT `TabMessage`)
**Discovery**: User manually tested - `TabMessage` doesn't show, `TabNewMailMessage` works

**Dropdown Pattern**: MUST use callbacks, NOT dynamic content
```xml
<dropDown id="ddActionSelector"
          getItemCount="GetActionCount"
          getItemLabel="GetActionLabel"
          getItemID="GetActionID"
          getSelectedItemIndex="GetSelectedActionIndex"
          onAction="OnActionSelected" />
```

**Why**: Using `getContent` with dynamic XML broke ribbon completely - ribbon stopped loading.

### COM Visibility
**Required**: `[ComVisible(true)]` on `InspectorComposeRibbon` class
**Required**: `IRibbonExtensibility` interface implementation

---

## 8. ERRORS ENCOUNTERED & FIXED

### Error 1: CS1061 - IRibbonUI.Context doesn't exist
**Fix**: Changed to `Globals.ThisAddIn.Application.ActiveInspector()`

### Error 2: Ribbon not appearing
**Root Cause**: Used `TabMessage` instead of `TabNewMailMessage`
**Fix**: User manually tested tabs, confirmed `TabNewMailMessage` works

### Error 3: _mailItem was null when pressing buttons
**Root Cause**: Cached reference became stale
**Fix**: Created `GetCurrentMailItem()` helper that gets fresh reference each time

### Error 4: Ribbon disappeared with nested box layout
**Root Cause**: Nested `<box>` elements broke ribbon XML parsing
**Fix**: Simplified to flat layout without nested containers

### Error 5: Ribbon disappeared with dynamic menu
**Root Cause**: Used `getContent="GetActionMenuContent"` for menu
**Fix**: Switched to `dropDown` with callback pattern

### Error 6: CS7036 - Missing packageContext parameter
**Root Cause**: `GroupActions()` call missing required parameters
**Fix**: Added `actionPane.CurrentPackageContext` and `CurrentProjectContext`

---

## 9. NEXT STEPS FOR NEW SESSION

### Immediate Action Required
1. **Rebuild with debugging code** that was just added but not yet tested:
```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "C:\Users\wally\source\repos\FastPMHelperAddin\FastPMHelperAddin.sln" /t:Rebuild /p:Configuration=Debug
```

2. **Reproduce the bug** and capture debug output showing:
   - How many times `CreateRibbonExtensibilityObject` is called
   - Which ribbon instance IDs are created (`Ribbon-1`, `Ribbon-2`, etc.)
   - Which instance `GetSelectedActionIndex` is called on for each window
   - What `mail.Subject` and `mail.EntryID` each callback sees
   - Whether `LoadDeferredData()` finds existing data or null

3. **Analyze debug output** to determine:
   - Is Outlook reusing the same ribbon instance for multiple windows?
   - Are callbacks executing on the correct instance for each window?
   - Is UserProperty data actually clean for new compose windows?

### Potential Solutions Based on Debug Output

**If Outlook Reuses Instances**:
Store state in a static dictionary keyed by Inspector/MailItem ID:
```csharp
private static Dictionary<string, RibbonState> _stateByMail = new Dictionary<string, RibbonState>();
```

**If GetCurrentMailItem Returns Wrong Reference**:
Use `ribbonUI.Context` or pass Inspector reference through InspectorWrapper

**If UserProperty Data Persists Unexpectedly**:
Investigate why new compose windows have existing UserProperty - may need to clear on Inspector close

### Alternative Approach to Consider
**Store state in InspectorWrapper instead of Ribbon instance**:
```csharp
// InspectorWrapper.cs
public class InspectorWrapper
{
    public ActionItem SelectedAction { get; set; }
    public DeferredActionData CurrentData { get; set; }
    // Ribbon retrieves state from wrapper, not its own fields
}
```

This ensures each Inspector has isolated state regardless of ribbon instance reuse.

---

## 10. VERIFICATION TESTS

### Test 1: Inline Compose (Sidebar Mode) ✅
**Status**: PASSING
1. Select email in Explorer → Click "Reply" inline
2. Sidebar enters compose mode, shows action buttons
3. Click "Create New" in Sidebar
4. Send email → Action executes

### Test 2: Pop-out Compose (Ribbon Mode) ✅
**Status**: PASSING (first window only)
1. Click "Reply" to open Inspector window
2. Inspector shows Ribbon with "Action Tracking" group
3. Sidebar returns to normal mode
4. Click "Create Action" toggle → saves UserProperty
5. Send email → Action executes

### Test 3: THE BUG FIX - Inline → Pop-out Transition ✅
**Status**: PASSING
1. Reply inline → Sidebar enters compose mode
2. Schedule action via Sidebar
3. Click "Pop Out"
4. **VERIFIED**: Sidebar immediately releases (no longer stuck!)
5. Inspector Ribbon shows scheduled action

### Test 4: Multiple Pop-out Windows ❌
**Status**: FAILING - This is the current bug
1. Open Inspector for Email A → Schedule "Update" for Action 123
2. Send Email A → Action executes ✅
3. Open Inspector for Email B
4. **BUG**: Shows Action 123 + "Update" toggle pressed
5. **EXPECTED**: Clean state, auto-select first linked action for Email B

---

## 11. USER FEEDBACK HISTORY

1. Initial request: "Implement the following plan: # Implementation Plan: Hybrid Architecture..."
2. After sidebar fix: "Great and I deleted tab message and it still shows..."
3. After dropdown added: "Great it is working now but drop down is not selectable"
4. After grouping added: "it does not have the linked action selected and it does not group them..."
5. **First bug report**: "Great that worked the first time but then it got stuck on that. So it had that linked action, + I ticked update, hit send, action updated correctly. But then every new draft I created (no matter what the actual linked action was) had that original linked action AND update already preselected"
6. After stateless attempt #1: "It is still persisting both the action AND the state."
7. After stateless attempt #2: "It is still doing the same thing"
8. Final message before handoff: "I want to start a new session. Write a detailed handoff.md..."

---

## 12. IMPORTANT CONTEXT

### UserProperty Naming
- **Prompt.txt specified**: `"Deferred_Execution_Data"`
- **Codebase uses**: `"FastPMDeferredAction"`
- **Both Sidebar and Ribbon use**: `"FastPMDeferredAction"` (ensures shared state)

### DeferredActionData Model
```csharp
public class DeferredActionData
{
    public string Mode { get; set; }        // "Create", "Update", "Close", "Reopen", "CreateMultiple"
    public int? ActionID { get; set; }      // Linked action ID (for Update/Close/Reopen)
    public string ManualTitle { get; set; }
    public string ManualAssignee { get; set; }
}
```

### ActionGroupingService
Groups actions into 4 categories:
- **Linked**: Actions directly linked to email's MessageID or InReplyTo
- **Package**: Actions in detected package context
- **Project**: Actions in detected project context
- **Other**: All remaining open actions

Dropdown labels use prefixes: `[LINKED]`, `[PKG: name]`, `[PRJ: name]`, `[OTHER]`

---

## 13. KNOWLEDGE BASE

### What Works in VSTO Ribbons
✅ Callback pattern for dropdowns (`getItemCount`, `getItemLabel`)
✅ `TabNewMailMessage` for compose windows
✅ Flat layout without nested containers
✅ `getPressed` callbacks for toggle state
✅ `_ribbon.Invalidate()` to refresh controls

### What Doesn't Work in VSTO Ribbons
❌ `getContent` with dynamic XML content
❌ Nested `<box>` elements
❌ `TabMessage` for compose windows
❌ `ribbonUI.Context` to get Inspector (doesn't exist)

### Important Patterns
- **Always use `GetCurrent*()` helpers** - never trust cached COM references
- **Old-style .csproj** - must manually add `<Compile>` and `<EmbeddedResource>` entries
- **MSBuild, not dotnet build** - VSTO projects require full MSBuild path
- **COM cleanup** - must release COM objects in finally blocks

---

## 14. DEBUG OUTPUT LOCATIONS

**Primary**: `C:\Users\wally\source\repos\FastPMHelperAddin\FastPMHelperAddin\Error.txt`

**Key Debug Lines Added** (not yet captured in output):
- `=== CreateRibbonExtensibilityObject CALLED at HH:mm:ss.fff ===`
- `*** InspectorComposeRibbon Constructor: Ribbon-N ***`
- `║ Ribbon_Load START: Ribbon-N at HH:mm:ss.fff ║`
- `[Ribbon-N] ⚠️ FOUND EXISTING DEFERRED DATA:` (indicates UserProperty found)
- `[Ribbon-N] ✓ No existing deferred data - CLEAN STATE` (indicates no UserProperty)
- `[Ribbon-N] ▶ GetSelectedActionIndex CALLED`
- `[Ribbon-N] GetUpdateActionPressed: Subject='...', Mode='...', returning TRUE/FALSE`

---

## 15. SUCCESS CRITERIA

The feature will be considered **COMPLETE** when:

✅ Sidebar releases immediately on pop-out (DONE)
✅ Inspector Ribbon loads in compose windows (DONE)
✅ Action dropdown shows grouped actions (DONE)
✅ Toggle buttons save to UserProperty (DONE)
✅ Actions execute on send (DONE)
❌ **Each compose window has isolated state** (FAILING)
❌ **No state carries over between different compose windows** (FAILING)

---

## 16. CONTACT & REFERENCES

**Codebase**: `C:\Users\wally\source\repos\FastPMHelperAddin\`
**Plan File**: `C:\Users\wally\.claude\plans\shimmering-scribbling-beaver.md`
**Build Command**: See section 7 - Technical Details
**Previous Session Transcript**: `C:\Users\wally\.claude\projects\C--Users-wally-source-repos-FastPMHelperAddin\5b82acb4-e4d8-4b04-96ce-892996dc9641.jsonl`

---

**End of Handoff Document**
**Priority**: 🔴 HIGH - Feature unusable until state persistence bug resolved
**Estimated Remaining Work**: 2-4 hours debugging + fix implementation
