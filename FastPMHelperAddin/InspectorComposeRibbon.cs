using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using FastPMHelperAddin.Models;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin
{
    /// <summary>
    /// Custom Ribbon for Inspector compose windows.
    /// Provides action tracking controls (Create, Update, Close, Reopen) for popped-out compose emails.
    /// </summary>
    [ComVisible(true)]
    public class InspectorComposeRibbon : Office.IRibbonExtensibility
    {
        private static int _instanceCounter = 0;
        private readonly int _instanceId;
        public string InstanceID => $"Ribbon-{_instanceId}";

        private Office.IRibbonUI _ribbon;
        private Outlook.Inspector _inspector;
        private Outlook.MailItem _mailItem;
        private DeferredActionData _currentDeferredData;
        private ActionItem _selectedAction;
        private List<ActionItem> _dropdownActions = new List<ActionItem>();
        private List<string> _dropdownLabels = new List<string>(); // Stores labels with section prefixes
        private int _linkedActionsCount = 0; // Track number of linked actions for auto-selection

        private const string DEFERRED_PROPERTY_NAME = "FastPMDeferredAction"; // Must match ProjectActionPane

        public InspectorComposeRibbon()
        {
            _instanceId = System.Threading.Interlocked.Increment(ref _instanceCounter);
            System.Diagnostics.Debug.WriteLine($"*** InspectorComposeRibbon Constructor: {InstanceID} ***");
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            // Only show ribbon for compose mail windows
            if (ribbonID == "Microsoft.Outlook.Mail.Compose")
            {
                return GetResourceText("FastPMHelperAddin.InspectorComposeRibbon.xml");
            }
            return null;
        }

        #endregion

        #region Ribbon Callbacks

        /// <summary>
        /// Called when the Ribbon loads. Stores the ribbon reference and initializes state.
        /// </summary>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            System.Diagnostics.Debug.WriteLine($"");
            System.Diagnostics.Debug.WriteLine($"╔═══════════════════════════════════════════════════════════════════");
            System.Diagnostics.Debug.WriteLine($"║ Ribbon_Load START: {InstanceID} at {DateTime.Now:HH:mm:ss.fff}");
            System.Diagnostics.Debug.WriteLine($"╚═══════════════════════════════════════════════════════════════════");

            _ribbon = ribbonUI;

            try
            {
                // IMPORTANT: Reset all state for this new ribbon instance
                _selectedAction = null;
                _currentDeferredData = null;
                _dropdownActions.Clear();
                _dropdownLabels.Clear();
                _linkedActionsCount = 0;

                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] State reset complete");

                // Get the Inspector that this Ribbon is being loaded for
                _inspector = Globals.ThisAddIn.Application.ActiveInspector();

                if (_inspector != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Inspector: {_inspector.GetHashCode()}");

                    // Get and cache the MailItem
                    _mailItem = GetCurrentMailItem();

                    if (_mailItem != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] MailItem:");
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   Subject: {_mailItem.Subject}");
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   EntryID: {_mailItem.EntryID ?? "(null - not saved yet)"}");

                        // Load any existing deferred action data from this email
                        _currentDeferredData = LoadDeferredData(_mailItem);

                        if (_currentDeferredData != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ⚠️ FOUND EXISTING DEFERRED DATA:");
                            System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   Mode: {_currentDeferredData.Mode}");
                            System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   ActionID: {_currentDeferredData.ActionID}");

                            // If there's a deferred action with an ActionID, find and set it as selected
                            if (_currentDeferredData.ActionID.HasValue)
                            {
                                var actionPane = Globals.ThisAddIn.GetActionPane();
                                if (actionPane?.OpenActions != null)
                                {
                                    _selectedAction = actionPane.OpenActions.FirstOrDefault(a => a.Id == _currentDeferredData.ActionID.Value);
                                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   Restored selected action: {_selectedAction?.Title ?? "NOT FOUND"}");
                                }
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ✓ No existing deferred data - CLEAN STATE");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ⚠️ WARNING: Could not get MailItem");
                    }

                    // Register this ribbon with the InspectorWrapper
                    var wrapper = Globals.ThisAddIn.GetInspectorWrapper(_inspector);
                    if (wrapper != null)
                    {
                        wrapper.Ribbon = this;
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Registered with InspectorWrapper");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ⚠️ WARNING: No active Inspector");
                }

                System.Diagnostics.Debug.WriteLine($"");
                System.Diagnostics.Debug.WriteLine($"╔═══════════════════════════════════════════════════════════════════");
                System.Diagnostics.Debug.WriteLine($"║ Ribbon_Load END: {InstanceID}");
                System.Diagnostics.Debug.WriteLine($"╚═══════════════════════════════════════════════════════════════════");
                System.Diagnostics.Debug.WriteLine($"");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ❌ ERROR in Ribbon_Load: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Stack: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Gets the current MailItem from the Inspector that triggered the callback.
        /// Uses the standard VSTO pattern: control.Context contains the Inspector for ribbon callbacks.
        /// </summary>
        private Outlook.MailItem GetCurrentMailItem(Office.IRibbonControl control)
        {
            try
            {
                // STANDARD VSTO PATTERN: Use control.Context to get the Inspector that triggered this callback
                // This works even when multiple Inspector windows are open
                var inspector = control.Context as Outlook.Inspector;
                if (inspector != null && inspector.CurrentItem is Outlook.MailItem mailItem)
                {
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetCurrentMailItem from control.Context: Subject='{mailItem.Subject}', EntryID='{mailItem.EntryID ?? "(null)"}'");
                    return mailItem;
                }

                // Fallback: try ActiveInspector if control.Context is null (shouldn't happen for Inspector ribbons)
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ⚠️ WARNING: control.Context is not an Inspector (type: {control.Context?.GetType().Name ?? "null"}), falling back to ActiveInspector");
                var activeInspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (activeInspector != null && activeInspector.CurrentItem is Outlook.MailItem activeMailItem)
                {
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Fallback succeeded: Subject='{activeMailItem.Subject}'");
                    return activeMailItem;
                }

                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ❌ ERROR: Could not get current MailItem from control.Context or ActiveInspector");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Error getting current MailItem: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets the current MailItem from the cached inspector (used in Ribbon_Load).
        /// This overload is kept for backwards compatibility with code that doesn't have access to control.
        /// </summary>
        private Outlook.MailItem GetCurrentMailItem()
        {
            try
            {
                // Try to get from cached inspector first
                if (_inspector != null && _inspector.CurrentItem is Outlook.MailItem mailItem)
                {
                    return mailItem;
                }

                // Fallback to active inspector
                var activeInspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (activeInspector != null && activeInspector.CurrentItem is Outlook.MailItem activeMailItem)
                {
                    return activeMailItem;
                }

                System.Diagnostics.Debug.WriteLine("WARNING: Could not get current MailItem");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting current MailItem: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Returns whether the "Create Action" toggle is pressed
        /// </summary>
        public bool GetCreateActionPressed(Office.IRibbonControl control)
        {
            var mail = GetCurrentMailItem(control);
            if (mail == null) return false;

            var data = LoadDeferredData(mail);
            return data?.Mode == "Create";
        }

        /// <summary>
        /// Handles the "Create Action" toggle button click
        /// </summary>
        public void OnCreateActionToggle(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"OnCreateActionToggle: pressed={pressed}");

                // Refresh the mail item reference
                var mail = GetCurrentMailItem(control);
                if (mail == null)
                {
                    System.Diagnostics.Debug.WriteLine("  ERROR: Could not get current MailItem");
                    return;
                }

                if (pressed)
                {
                    // Clear any existing schedule (mutual exclusivity)
                    if (_currentDeferredData != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"  Clearing existing {_currentDeferredData.Mode} schedule");
                        ClearDeferredData(mail);
                    }

                    // Schedule Create action
                    var data = new DeferredActionData { Mode = "Create" };
                    SaveDeferredData(mail, data);
                    _currentDeferredData = data;
                    System.Diagnostics.Debug.WriteLine("  Create action scheduled");
                }
                else
                {
                    // Cancel schedule
                    ClearDeferredData(mail);
                    _currentDeferredData = null;
                    _selectedAction = null; // Clear selection when canceling
                    System.Diagnostics.Debug.WriteLine("  Create action schedule canceled");
                }

                InvalidateRibbon();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnCreateActionToggle: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns whether the "Create Multiple" toggle is pressed
        /// </summary>
        public bool GetCreateMultipleActionPressed(Office.IRibbonControl control)
        {
            var mail = GetCurrentMailItem(control);
            if (mail == null) return false;

            var data = LoadDeferredData(mail);
            return data?.Mode == "CreateMultiple";
        }

        /// <summary>
        /// Handles the "Create Multiple" toggle button click
        /// </summary>
        public void OnCreateMultipleActionToggle(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"OnCreateMultipleActionToggle: pressed={pressed}");

                // Refresh the mail item reference
                var mail = GetCurrentMailItem(control);
                if (mail == null)
                {
                    System.Diagnostics.Debug.WriteLine("  ERROR: Could not get current MailItem");
                    return;
                }

                if (pressed)
                {
                    // Clear any existing schedule (mutual exclusivity)
                    if (_currentDeferredData != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"  Clearing existing {_currentDeferredData.Mode} schedule");
                        ClearDeferredData(mail);
                    }

                    // Schedule CreateMultiple action
                    var data = new DeferredActionData { Mode = "CreateMultiple" };
                    SaveDeferredData(mail, data);
                    _currentDeferredData = data;
                    System.Diagnostics.Debug.WriteLine("  CreateMultiple action scheduled");
                }
                else
                {
                    // Cancel schedule
                    ClearDeferredData(mail);
                    _currentDeferredData = null;
                    _selectedAction = null; // Clear selection when canceling
                    System.Diagnostics.Debug.WriteLine("  CreateMultiple action schedule canceled");
                }

                InvalidateRibbon();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnCreateMultipleActionToggle: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns whether the "Update Action" toggle is pressed
        /// </summary>
        public bool GetUpdateActionPressed(Office.IRibbonControl control)
        {
            var mail = GetCurrentMailItem(control);
            if (mail == null)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetUpdateActionPressed: mail=null, returning FALSE");
                return false;
            }

            var data = LoadDeferredData(mail);
            bool isPressed = data?.Mode == "Update";
            System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetUpdateActionPressed: Subject='{mail.Subject}', Mode='{data?.Mode}', returning {isPressed}");
            return isPressed;
        }

        /// <summary>
        /// Handles the "Update Action" toggle button click
        /// </summary>
        public void OnUpdateActionToggle(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"OnUpdateActionToggle: pressed={pressed}");

                // Refresh the mail item reference
                var mail = GetCurrentMailItem(control);
                if (mail == null)
                {
                    System.Diagnostics.Debug.WriteLine("  ERROR: Could not get current MailItem");
                    return;
                }

                if (pressed)
                {
                    // Clear any existing schedule (mutual exclusivity)
                    if (_currentDeferredData != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"  Clearing existing {_currentDeferredData.Mode} schedule");
                        ClearDeferredData(mail);
                    }

                    // Schedule Update action with selected action ID
                    var data = new DeferredActionData
                    {
                        Mode = "Update",
                        ActionID = _selectedAction?.Id
                    };
                    SaveDeferredData(mail, data);
                    _currentDeferredData = data;
                    System.Diagnostics.Debug.WriteLine($"  Update action scheduled for ActionID: {_selectedAction?.Id}");
                }
                else
                {
                    // Cancel schedule
                    ClearDeferredData(mail);
                    _currentDeferredData = null;
                    _selectedAction = null; // Clear selection when canceling
                    System.Diagnostics.Debug.WriteLine("  Update action schedule canceled");
                }

                InvalidateRibbon();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnUpdateActionToggle: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns whether the "Close Action" toggle is pressed
        /// </summary>
        public bool GetCloseActionPressed(Office.IRibbonControl control)
        {
            var mail = GetCurrentMailItem(control);
            if (mail == null) return false;

            var data = LoadDeferredData(mail);
            return data?.Mode == "Close";
        }

        /// <summary>
        /// Handles the "Close Action" toggle button click
        /// </summary>
        public void OnCloseActionToggle(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"OnCloseActionToggle: pressed={pressed}");

                // Refresh the mail item reference
                var mail = GetCurrentMailItem(control);
                if (mail == null)
                {
                    System.Diagnostics.Debug.WriteLine("  ERROR: Could not get current MailItem");
                    return;
                }

                if (pressed)
                {
                    // Clear any existing schedule (mutual exclusivity)
                    if (_currentDeferredData != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"  Clearing existing {_currentDeferredData.Mode} schedule");
                        ClearDeferredData(mail);
                    }

                    // Schedule Close action with selected action ID
                    var data = new DeferredActionData
                    {
                        Mode = "Close",
                        ActionID = _selectedAction?.Id
                    };
                    SaveDeferredData(mail, data);
                    _currentDeferredData = data;
                    System.Diagnostics.Debug.WriteLine($"  Close action scheduled for ActionID: {_selectedAction?.Id}");
                }
                else
                {
                    // Cancel schedule
                    ClearDeferredData(mail);
                    _currentDeferredData = null;
                    _selectedAction = null; // Clear selection when canceling
                    System.Diagnostics.Debug.WriteLine("  Close action schedule canceled");
                }

                InvalidateRibbon();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnCloseActionToggle: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns whether the "Reopen Action" toggle is pressed
        /// </summary>
        public bool GetReopenActionPressed(Office.IRibbonControl control)
        {
            var mail = GetCurrentMailItem(control);
            if (mail == null) return false;

            var data = LoadDeferredData(mail);
            return data?.Mode == "Reopen";
        }

        /// <summary>
        /// Handles the "Reopen Action" toggle button click
        /// </summary>
        public void OnReopenActionToggle(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"OnReopenActionToggle: pressed={pressed}");

                // Refresh the mail item reference
                var mail = GetCurrentMailItem(control);
                if (mail == null)
                {
                    System.Diagnostics.Debug.WriteLine("  ERROR: Could not get current MailItem");
                    return;
                }

                if (pressed)
                {
                    // Clear any existing schedule (mutual exclusivity)
                    if (_currentDeferredData != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"  Clearing existing {_currentDeferredData.Mode} schedule");
                        ClearDeferredData(mail);
                    }

                    // Schedule Reopen action with selected action ID
                    var data = new DeferredActionData
                    {
                        Mode = "Reopen",
                        ActionID = _selectedAction?.Id
                    };
                    SaveDeferredData(mail, data);
                    _currentDeferredData = data;
                    System.Diagnostics.Debug.WriteLine($"  Reopen action scheduled for ActionID: {_selectedAction?.Id}");
                }
                else
                {
                    // Cancel schedule
                    ClearDeferredData(mail);
                    _currentDeferredData = null;
                    _selectedAction = null; // Clear selection when canceling
                    System.Diagnostics.Debug.WriteLine("  Reopen action schedule canceled");
                }

                InvalidateRibbon();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnReopenActionToggle: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns the status label text
        /// </summary>
        public string GetStatusLabel(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetStatusLabel CALLED");

                // Load fresh data from the current Inspector's mail item
                var mail = GetCurrentMailItem(control);
                if (mail == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetStatusLabel: mail is null");
                    return "Status: No mail item";
                }

                var data = LoadDeferredData(mail);
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetStatusLabel: Mode='{data?.Mode}', ActionID={data?.ActionID}");

                if (data == null || string.IsNullOrEmpty(data.Mode))
                {
                    // No scheduled action - check if there's a selected action for display purposes
                    if (_selectedAction != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetStatusLabel: Returning selected action ID: {_selectedAction.Id}");
                        return $"Selected: ID {_selectedAction.Id}";
                    }
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetStatusLabel: No action scheduled");
                    return "Status: No action scheduled";
                }

                // Show the scheduled action from this email's UserProperty
                // Keep it short to prevent ribbon layout issues
                string modeText = data.Mode;
                if (data.ActionID.HasValue)
                {
                    modeText += $" ID: {data.ActionID}";
                }

                string result = $"Scheduled: {modeText}";
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetStatusLabel: Returning '{result}'");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ERROR in GetStatusLabel: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Stack: {ex.StackTrace}");
                return "Status: Error";
            }
        }

        /// <summary>
        /// Populates the dropdown actions list from the action pane.
        /// Uses captured grouping from InspectorWrapper (set at pop-out time) for accurate results.
        /// </summary>
        private void PopulateDropdownActions(Office.IRibbonControl control)
        {
            try
            {
                _dropdownActions.Clear();
                _dropdownLabels.Clear();
                _linkedActionsCount = 0;

                var actionPane = Globals.ThisAddIn.GetActionPane();
                if (actionPane == null) return;

                var openActions = actionPane.OpenActions;
                if (openActions == null || openActions.Count == 0) return;

                // Try to get the InspectorWrapper for this Inspector to use captured grouping
                var inspector = control.Context as Outlook.Inspector;
                Services.ActionGroupingResult grouping = null;

                if (inspector != null)
                {
                    var wrapper = Globals.ThisAddIn.GetInspectorWrapper(inspector);
                    if (wrapper?.CapturedGrouping != null)
                    {
                        grouping = wrapper.CapturedGrouping;
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Using CAPTURED grouping from InspectorWrapper");
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   Captured: {grouping.LinkedActions.Count} linked, {grouping.PackageActions.Count} package, {grouping.ProjectActions.Count} project, {grouping.OtherActions.Count} other");
                    }
                }

                // If no captured grouping, calculate it (fallback for edge cases)
                if (grouping == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] No captured grouping - calculating from scratch");

                    var mail = GetCurrentMailItem(control);
                    if (mail == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] WARNING: GetCurrentMailItem returned null - using fallback");
                        foreach (var action in openActions.Take(50))
                        {
                            _dropdownActions.Add(action);
                            _dropdownLabels.Add(action.Title);
                        }
                        return;
                    }

                    var emailProps = Models.EmailProperties.ExtractFrom(mail);
                    if (emailProps == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] WARNING: Failed to extract email properties - using fallback");
                        foreach (var action in openActions.Take(50))
                        {
                            _dropdownActions.Add(action);
                            _dropdownLabels.Add(action.Title);
                        }
                        return;
                    }

                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] PopulateDropdownActions for email: '{emailProps.Subject}'");

                    var groupingService = actionPane.GroupingService;
                    if (groupingService != null)
                    {
                        // Use sidebar's ConversationId as fallback for drafts
                        string conversationIdToUse = emailProps.ConversationId;
                        if (string.IsNullOrEmpty(conversationIdToUse) && actionPane.CurrentEmailProperties != null)
                        {
                            conversationIdToUse = actionPane.CurrentEmailProperties.ConversationId;
                        }

                        grouping = groupingService.GroupActions(
                            openActions,
                            emailProps.InternetMessageId,
                            emailProps.InReplyToId,
                            conversationIdToUse,
                            actionPane.CurrentPackageContext ?? "",
                            actionPane.CurrentProjectContext ?? ""
                        );
                    }
                }

                // Use the grouping results to populate dropdown
                if (grouping != null)
                {
                    // Add Linked Actions
                    if (grouping.LinkedActions.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   Linked actions:");
                        foreach (var action in grouping.LinkedActions)
                        {
                            System.Diagnostics.Debug.WriteLine($"[{InstanceID}]     - ID {action.Id}: {action.Title}");
                            _dropdownActions.Add(action);
                            _dropdownLabels.Add($"[LINKED] {action.Title}");
                        }
                        _linkedActionsCount = grouping.LinkedActions.Count;
                    }

                    // Add Package Actions
                    if (grouping.PackageActions.Count > 0)
                    {
                        string packagePrefix = string.IsNullOrEmpty(grouping.DetectedPackage)
                            ? "[PACKAGE]"
                            : $"[PKG: {grouping.DetectedPackage}]";

                        foreach (var action in grouping.PackageActions)
                        {
                            _dropdownActions.Add(action);
                            _dropdownLabels.Add($"{packagePrefix} {action.Title}");
                        }
                    }

                    // Add Project Actions
                    if (grouping.ProjectActions.Count > 0)
                    {
                        string projectPrefix = string.IsNullOrEmpty(grouping.DetectedProject)
                            ? "[PROJECT]"
                            : $"[PRJ: {grouping.DetectedProject}]";

                        foreach (var action in grouping.ProjectActions)
                        {
                            _dropdownActions.Add(action);
                            _dropdownLabels.Add($"{projectPrefix} {action.Title}");
                        }
                    }

                    // Add Other Actions
                    if (grouping.OtherActions.Count > 0)
                    {
                        foreach (var action in grouping.OtherActions)
                        {
                            _dropdownActions.Add(action);
                            _dropdownLabels.Add($"[OTHER] {action.Title}");
                        }
                    }
                }
                else
                {
                    // Fallback: use all open actions without grouping
                    foreach (var action in openActions.Take(50))
                    {
                        _dropdownActions.Add(action);
                        _dropdownLabels.Add(action.Title);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Error in PopulateDropdownActions: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns the number of items in the dropdown
        /// </summary>
        public int GetActionCount(Office.IRibbonControl control)
        {
            System.Diagnostics.Debug.WriteLine($"[{InstanceID}] GetActionCount CALLED");
            PopulateDropdownActions(control);
            System.Diagnostics.Debug.WriteLine($"[{InstanceID}] PopulateDropdownActions complete - {_dropdownActions.Count} actions");
            return _dropdownActions.Count;
        }

        /// <summary>
        /// Returns the ID for the dropdown item at the specified index
        /// </summary>
        public string GetActionID(Office.IRibbonControl control, int index)
        {
            if (index >= 0 && index < _dropdownActions.Count)
            {
                return _dropdownActions[index].Id.ToString();
            }
            return "0";
        }

        /// <summary>
        /// Returns the label for the dropdown item at the specified index
        /// </summary>
        public string GetActionLabel(Office.IRibbonControl control, int index)
        {
            if (index >= 0 && index < _dropdownLabels.Count)
            {
                return _dropdownLabels[index];
            }
            return "(No actions)";
        }

        /// <summary>
        /// Returns the selected item index (auto-selects first linked action)
        /// </summary>
        public int GetSelectedActionIndex(Office.IRibbonControl control)
        {
            System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ▶ GetSelectedActionIndex CALLED");

            // Get fresh deferred data from current mail item
            var mail = GetCurrentMailItem(control);
            if (mail != null)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   MailItem Subject: '{mail.Subject}'");

                var data = LoadDeferredData(mail);
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   Loaded data: Mode='{data?.Mode}', ActionID={data?.ActionID}");

                // If there's saved deferred data with an ActionID, select that action
                if (data?.ActionID.HasValue == true)
                {
                    for (int i = 0; i < _dropdownActions.Count; i++)
                    {
                        if (_dropdownActions[i].Id == data.ActionID.Value)
                        {
                            System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   ✓ RETURNING index {i} from SAVED DATA: {_dropdownActions[i].Title}");
                            return i;
                        }
                    }
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   ⚠️ WARNING: ActionID {data.ActionID.Value} not found in dropdown");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   ⚠️ WARNING: mail is null");
            }

            // If user has manually selected an action (but no saved data), find its index
            if (_selectedAction != null)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   Checking _selectedAction: {_selectedAction.Title} (ID: {_selectedAction.Id})");
                for (int i = 0; i < _dropdownActions.Count; i++)
                {
                    if (_dropdownActions[i].Id == _selectedAction.Id)
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   ✓ RETURNING index {i} from _selectedAction field");
                        return i;
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   ⚠️ WARNING: _selectedAction not found in dropdown");
            }

            // Auto-select first linked action if available and no saved data
            if (_linkedActionsCount > 0 && _dropdownActions.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   ✓ RETURNING index 0 (auto-select first linked): {_dropdownActions[0].Title}");
                return 0;
            }

            // No selection
            System.Diagnostics.Debug.WriteLine($"[{InstanceID}]   ✓ RETURNING -1 (no selection)");
            return -1;
        }

        /// <summary>
        /// Handles action selection from the dropdown
        /// </summary>
        public void OnActionSelected(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] OnActionSelected: selectedIndex={selectedIndex}, selectedId={selectedId}");

                if (selectedIndex >= 0 && selectedIndex < _dropdownActions.Count)
                {
                    _selectedAction = _dropdownActions[selectedIndex];
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Selected action: {_selectedAction?.Title ?? "null"}");

                    // Load fresh deferred data from the current Inspector's mail item
                    var mail = GetCurrentMailItem(control);
                    if (mail == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ERROR: Could not get current MailItem in OnActionSelected");
                        return;
                    }

                    var currentData = LoadDeferredData(mail);

                    // Update deferred data if Update/Close/Reopen is scheduled
                    if (currentData != null &&
                        (currentData.Mode == "Update" || currentData.Mode == "Close" || currentData.Mode == "Reopen"))
                    {
                        currentData.ActionID = _selectedAction.Id;
                        SaveDeferredData(mail, currentData);
                        _currentDeferredData = currentData; // Update instance field for consistency
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Updated {currentData.Mode} action with ActionID: {_selectedAction.Id}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[{InstanceID}] No scheduled action to update (Mode: {currentData?.Mode ?? "null"})");
                    }

                    InvalidateRibbon();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[{InstanceID}] WARNING: Invalid selectedIndex {selectedIndex} (dropdown has {_dropdownActions.Count} items)");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] ERROR in OnActionSelected: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[{InstanceID}] Stack: {ex.StackTrace}");
            }
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Invalidates (refreshes) all ribbon controls
        /// </summary>
        public void InvalidateRibbon()
        {
            _ribbon?.Invalidate();
        }

        /// <summary>
        /// Saves deferred action data to the mail item's UserProperties
        /// </summary>
        private void SaveDeferredData(Outlook.MailItem mail, DeferredActionData data)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== SaveDeferredData START ===");
                System.Diagnostics.Debug.WriteLine($"Mail subject: {mail.Subject}");
                System.Diagnostics.Debug.WriteLine($"Data Mode: {data.Mode}, ActionID: {data.ActionID}");

                var json = JsonSerializer.Serialize(data);
                System.Diagnostics.Debug.WriteLine($"Serialized JSON: {json}");

                var props = mail.UserProperties;
                var prop = props.Find(DEFERRED_PROPERTY_NAME);

                if (prop == null)
                {
                    prop = props.Add(DEFERRED_PROPERTY_NAME, Outlook.OlUserPropertyType.olText);
                    System.Diagnostics.Debug.WriteLine("  Created new UserProperty");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  Found existing UserProperty");
                }

                prop.Value = json;
                mail.Save();

                System.Diagnostics.Debug.WriteLine($"✓ Property saved and mail.Save() called");
                System.Diagnostics.Debug.WriteLine($"=== SaveDeferredData END ===");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR saving deferred data: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Loads deferred action data from the mail item's UserProperties
        /// </summary>
        private DeferredActionData LoadDeferredData(Outlook.MailItem mail)
        {
            try
            {
                var props = mail.UserProperties;
                var prop = props.Find(DEFERRED_PROPERTY_NAME);

                if (prop == null || string.IsNullOrEmpty(prop.Value?.ToString()))
                {
                    return null;
                }

                var json = prop.Value.ToString();
                return JsonSerializer.Deserialize<DeferredActionData>(json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR loading deferred data: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Clears deferred action data from the mail item's UserProperties
        /// </summary>
        private void ClearDeferredData(Outlook.MailItem mail)
        {
            try
            {
                var props = mail.UserProperties;
                var prop = props.Find(DEFERRED_PROPERTY_NAME);

                if (prop != null)
                {
                    prop.Delete();
                    mail.Save();
                    System.Diagnostics.Debug.WriteLine("UserProperty cleared and mail saved");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR clearing deferred data: {ex.Message}");
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
