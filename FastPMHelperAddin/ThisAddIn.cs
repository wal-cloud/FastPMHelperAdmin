using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms.Integration;
using FastPMHelperAddin.Models;
using FastPMHelperAddin.UI;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin
{
    public partial class ThisAddIn
    {
        private Outlook.Application _app;
        private Outlook.NameSpace _ns;
        private List<Outlook.Items> _monitoredInboxes = new List<Outlook.Items>();

        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private ProjectActionPane _actionPane;
        private Outlook.Explorer _currentExplorer;
        private Outlook.Inspectors _inspectors;
        private Dictionary<string, InspectorWrapper> _inspectorWrappers = new Dictionary<string, InspectorWrapper>();

        private const string RootPath = @"C:\MailPipeline";
        private const string TrackedEmailAccount = "wally.cloud@dynonobel.com";

        // Deferred action property name
        private const string DEFERRED_PROPERTY_NAME = "FastPMDeferredAction";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                _app = Globals.ThisAddIn.Application;
                _ns = _app.GetNamespace("MAPI");

                // Monitor Sent Items for ALL accounts (for deferred action execution)
                System.Diagnostics.Debug.WriteLine("=== Setting up Sent Items monitoring ===");
                foreach (Outlook.Store store in _ns.Stores)
                {
                    try
                    {
                        var rootFolder = store.GetRootFolder();
                        var sentFolder = GetSentItemsFolder(rootFolder);

                        if (sentFolder != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"Monitoring Sent Items: {sentFolder.Name} in store: {store.DisplayName}");
                            var items = sentFolder.Items;
                            items.ItemAdd += SentItems_ItemAdd;
                            _monitoredInboxes.Add(items); // Reuse the list to track all monitored folders
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error setting up Sent Items monitoring for store {store.DisplayName}: {ex.Message}");
                    }
                }

                // NEW: Initialize Custom TaskPane
                InitializeTaskPane();

                // NEW: Wire up Explorer SelectionChange event
                try
                {
                    _currentExplorer = _app.ActiveExplorer();
                    if (_currentExplorer != null)
                    {
                        System.Diagnostics.Debug.WriteLine("Hooking up Explorer SelectionChange event...");
                        _currentExplorer.SelectionChange += Explorer_SelectionChange;
                        System.Diagnostics.Debug.WriteLine($"Explorer SelectionChange event hooked successfully (Explorer: {_currentExplorer.GetHashCode()})");

                        // Hook InlineResponse event for inline compose detection
                        _currentExplorer.InlineResponse += Explorer_InlineResponse;
                        System.Diagnostics.Debug.WriteLine("Explorer InlineResponse event hooked successfully");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("WARNING: ActiveExplorer is null, cannot hook SelectionChange event");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error hooking SelectionChange event: {ex.Message}");
                }

                // NEW: Wire up Inspectors event for compose window detection
                try
                {
                    _inspectors = _app.Inspectors;
                    _inspectors.NewInspector += Inspectors_NewInspector;
                    System.Diagnostics.Debug.WriteLine("Inspectors.NewInspector event hooked successfully");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error hooking Inspectors event: {ex.Message}");
                }

                System.Diagnostics.Debug.WriteLine("MailPipeline add-in started successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting add-in: {ex.Message}");
            }
        }

        private Outlook.MAPIFolder GetInboxFolder(Outlook.MAPIFolder parentFolder)
        {
            try
            {
                // Try to get the default inbox first
                if (parentFolder.DefaultMessageClass == "IPM.Note" &&
                    parentFolder.Name.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
                {
                    return parentFolder;
                }

                // Search subfolders
                foreach (Outlook.MAPIFolder folder in parentFolder.Folders)
                {
                    if (folder.DefaultMessageClass == "IPM.Note" &&
                        folder.Name.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
                    {
                        return folder;
                    }

                    // Recursive search
                    var found = GetInboxFolder(folder);
                    if (found != null)
                        return found;
                }
            }
            catch { }

            return null;
        }

        private Outlook.MAPIFolder GetSentItemsFolder(Outlook.MAPIFolder parentFolder)
        {
            try
            {
                // Try to get the sent items folder
                if (parentFolder.DefaultMessageClass == "IPM.Note" &&
                    (parentFolder.Name.Equals("Sent Items", StringComparison.OrdinalIgnoreCase) ||
                     parentFolder.Name.Equals("Sent", StringComparison.OrdinalIgnoreCase)))
                {
                    return parentFolder;
                }

                // Search subfolders
                foreach (Outlook.MAPIFolder folder in parentFolder.Folders)
                {
                    if (folder.DefaultMessageClass == "IPM.Note" &&
                        (folder.Name.Equals("Sent Items", StringComparison.OrdinalIgnoreCase) ||
                         folder.Name.Equals("Sent", StringComparison.OrdinalIgnoreCase)))
                    {
                        return folder;
                    }

                    // Recursive search
                    var found = GetSentItemsFolder(folder);
                    if (found != null)
                        return found;
                }
            }
            catch { }

            return null;
        }

        private void InitializeTaskPane()
        {
            try
            {
                _actionPane = new ProjectActionPane();

                // Wrap WPF control in UserControl for VSTO compatibility
                var wrapper = new TaskPaneWrapper(_actionPane);

                // Create the task pane
                _taskPane = this.CustomTaskPanes.Add(wrapper, "Project Actions");
                _taskPane.Width = 320;
                _taskPane.Visible = true;

                System.Diagnostics.Debug.WriteLine("Task pane initialized successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error initializing task pane: {ex.Message}");
            }
        }

        private void Explorer_SelectionChange()
        {
            System.Diagnostics.Debug.WriteLine($"=== Explorer_SelectionChange EVENT FIRED === (Time: {DateTime.Now:HH:mm:ss.fff})");

            try
            {
                // Check if we need to exit compose mode (inline compose only)
                if (_actionPane != null && _actionPane.IsComposeMode)
                {
                    // Only handle inline compose exit (popup compose handled by Inspector.Close)
                    if (_actionPane.ComposeInspector == null)
                    {
                        System.Diagnostics.Debug.WriteLine("  In inline compose mode - checking if user clicked away");
                        // Inline compose mode - check if user clicked away
                        if (_currentExplorer.ActiveInlineResponse == null)
                        {
                            System.Diagnostics.Debug.WriteLine("  ActiveInlineResponse is null - exiting compose mode");
                            _actionPane.Dispatcher.Invoke(() =>
                            {
                                _actionPane.OnComposeItemDeactivated();
                            });
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("  ActiveInlineResponse still active - staying in compose mode");
                            // Don't process selection - we're still in compose mode
                            return;
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("  In popup compose mode - Inspector.Close will handle exit");
                        // Don't process selection - we're in popup compose mode
                        return;
                    }
                }

                if (_actionPane == null)
                {
                    System.Diagnostics.Debug.WriteLine("  _actionPane is null, returning");
                    return;
                }

                var explorer = _app.ActiveExplorer();
                System.Diagnostics.Debug.WriteLine($"  Current Explorer: {explorer?.GetHashCode()}, Stored Explorer: {_currentExplorer?.GetHashCode()}");
                System.Diagnostics.Debug.WriteLine($"  Selection count: {explorer.Selection.Count}");

                if (explorer.Selection.Count > 0)
                {
                    var selection = explorer.Selection[1]; // Outlook is 1-indexed
                    System.Diagnostics.Debug.WriteLine($"  Selection type: {selection?.GetType().Name}");

                    if (selection is Outlook.MailItem mail)
                    {
                        System.Diagnostics.Debug.WriteLine($"  Email selected: {mail.Subject}");
                        System.Diagnostics.Debug.WriteLine($"  Sent status: {mail.Sent}");

                        // NEW: Check if this is an unsent email (draft/compose mode)
                        if (!mail.Sent)
                        {
                            System.Diagnostics.Debug.WriteLine("  Detected UNSENT email (inline compose) - entering compose mode");

                            _actionPane.Dispatcher.Invoke(() =>
                            {
                                // Enter compose mode for inline editing (no Inspector window)
                                _actionPane.OnComposeItemActivated(mail, null);
                            });
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("  Regular sent/received email - normal mode");

                            // PERFORMANCE FIX: Extract properties on COM thread before marshaling to UI thread
                            // This avoids blocking the UI thread with synchronous Outlook PropertyAccessor calls
                            var sw = System.Diagnostics.Stopwatch.StartNew();
                            var emailProperties = EmailProperties.ExtractFrom(mail);
                            sw.Stop();
                            System.Diagnostics.Debug.WriteLine($"  Property extraction took {sw.ElapsedMilliseconds}ms on COM thread");

                            // Marshal to WPF UI thread (COM -> WPF threading)
                            _actionPane.Dispatcher.Invoke(() =>
                            {
                                _actionPane.OnEmailSelected(emailProperties);
                            });
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("  Selection is not a MailItem");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  No selection - clearing pane");
                    // No selection - clear the pane (pass null EmailProperties)
                    _actionPane.Dispatcher.Invoke(() =>
                    {
                        _actionPane.OnEmailSelected(null);
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SelectionChange: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        private void Explorer_InlineResponse(object item)
        {
            System.Diagnostics.Debug.WriteLine($"=== Explorer_InlineResponse EVENT FIRED === (Time: {DateTime.Now:HH:mm:ss.fff})");

            try
            {
                if (item is Outlook.MailItem draft)
                {
                    System.Diagnostics.Debug.WriteLine($"  Inline compose started for: {draft.Subject}");

                    // Race condition check: If already in compose mode with different draft
                    if (_actionPane.IsComposeMode && _actionPane.ComposeMail != null)
                    {
                        if (draft.EntryID != _actionPane.ComposeMail.EntryID)
                        {
                            System.Diagnostics.Debug.WriteLine($"  Different draft detected - exiting old compose mode first");
                            // Different draft - exit old, enter new
                            _actionPane.Dispatcher.Invoke(() =>
                            {
                                _actionPane.OnComposeItemDeactivated();
                            });
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"  Same draft - already in compose mode, ignoring");
                            // Same draft - already in compose mode, ignore
                            return;
                        }
                    }

                    // Inline compose started in reading pane
                    System.Diagnostics.Debug.WriteLine($"  Entering compose mode (inline)");
                    _actionPane.Dispatcher.Invoke(() =>
                    {
                        _actionPane.OnComposeItemActivated(draft, null);
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in InlineResponse: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            System.Diagnostics.Debug.WriteLine("=== Inspectors_NewInspector EVENT FIRED ===");

            try
            {
                if (_actionPane == null)
                {
                    System.Diagnostics.Debug.WriteLine("  _actionPane is null, returning");
                    return;
                }

                // Check if this is a compose window (unsent MailItem)
                if (inspector.CurrentItem is Outlook.MailItem mail)
                {
                    System.Diagnostics.Debug.WriteLine($"  Inspector contains MailItem: {mail.Subject ?? "(No Subject)"}");
                    System.Diagnostics.Debug.WriteLine($"  Sent status: {mail.Sent}");

                    // Only process unsent emails (compose mode)
                    if (!mail.Sent)
                    {
                        System.Diagnostics.Debug.WriteLine("  Detected compose window - creating InspectorWrapper");

                        // Create InspectorWrapper to manage this window with its Ribbon
                        var wrapper = new InspectorWrapper(inspector);
                        string key = GetInspectorKey(inspector);
                        _inspectorWrappers[key] = wrapper;

                        // CRITICAL FIX: Release Sidebar immediately (don't call OnComposeItemActivated!)
                        // The Ribbon will now control this Inspector window
                        _actionPane.Dispatcher.Invoke(() =>
                        {
                            if (_actionPane.IsComposeMode)
                            {
                                System.Diagnostics.Debug.WriteLine("  Sidebar was in compose mode - releasing to Ribbon");
                                _actionPane.OnComposeItemDeactivated();
                            }
                        });
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("  Email already sent - not a compose window");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"  Inspector does not contain MailItem (type: {inspector.CurrentItem?.GetType().Name ?? "null"})");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Inspectors_NewInspector: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                // Unhook Sent Items event handlers
                foreach (var items in _monitoredInboxes)
                {
                    if (items != null)
                        items.ItemAdd -= SentItems_ItemAdd;
                }

                // NEW: Unhook explorer events
                if (_currentExplorer != null)
                {
                    _currentExplorer.SelectionChange -= Explorer_SelectionChange;
                    _currentExplorer.InlineResponse -= Explorer_InlineResponse;
                    System.Diagnostics.Debug.WriteLine("Explorer events unhooked (SelectionChange, InlineResponse)");
                }

                // NEW: Unhook inspectors event
                if (_inspectors != null)
                {
                    _inspectors.NewInspector -= Inspectors_NewInspector;
                    System.Diagnostics.Debug.WriteLine("Inspectors.NewInspector event unhooked");
                }

                // Clean up InspectorWrapper dictionary
                _inspectorWrappers.Clear();

                System.Diagnostics.Debug.WriteLine("MailPipeline add-in shut down successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during shutdown: {ex.Message}");
            }
        }


        private async void SentItems_ItemAdd(object item)
        {
            System.Diagnostics.Debug.WriteLine("=== SentItems_ItemAdd EVENT FIRED ===");

            if (item is Outlook.MailItem mail)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"Processing sent mail: {mail.Subject}");

                    // Check for deferred action execution
                    var deferredData = LoadDeferredData(mail);
                    if (deferredData != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Found deferred action: {deferredData.Mode}");
                        await ExecuteDeferredActionAsync(mail, deferredData);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("No deferred action scheduled for this email");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in SentItems_ItemAdd: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Item is not a MailItem, type: {item?.GetType().Name ?? "null"}");
            }
        }


        public string GetSenderEmailAddress(Outlook.MailItem mail)
        {
            // Check the folder to determine if this is a sent or received email
            var folder = mail.Parent as Outlook.MAPIFolder;
            bool isSentFolder = folder != null &&
                               (folder.Name.Equals("Sent Items", StringComparison.OrdinalIgnoreCase) ||
                                folder.Name.Equals("Sent", StringComparison.OrdinalIgnoreCase));

            // For sent items, get the sending account
            if (isSentFolder && mail.SendUsingAccount != null)
                return mail.SendUsingAccount.SmtpAddress;

            // For inbox items or when SendUsingAccount is null, use the actual sender
            // Try SenderEmailAddress property
            if (!string.IsNullOrEmpty(mail.SenderEmailAddress))
            {
                // If it's an Exchange address, try to get SMTP
                if (mail.SenderEmailAddress.StartsWith("/"))
                {
                    try
                    {
                        var sender = mail.Sender;
                        if (sender != null)
                            return sender.GetExchangeUser()?.PrimarySmtpAddress ?? mail.SenderEmailAddress;
                    }
                    catch { }
                }
                return mail.SenderEmailAddress;
            }

            return string.Empty;
        }

        private string GetRecipientEmailAddress(Outlook.Recipient recipient)
        {
            try
            {
                if (recipient.AddressEntry?.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    recipient.AddressEntry?.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var exchangeUser = recipient.AddressEntry.GetExchangeUser();
                    if (exchangeUser != null)
                        return exchangeUser.PrimarySmtpAddress;
                }
                return recipient.Address;
            }
            catch
            {
                return recipient.Address;
            }
        }

        #region Deferred Action Execution

        /// <summary>
        /// Loads deferred action data from the mail item's UserProperties
        /// </summary>
        private DeferredActionData LoadDeferredData(Outlook.MailItem mail)
        {
            try
            {
                var props = mail.UserProperties;

                // Log all user properties for debugging
                System.Diagnostics.Debug.WriteLine($"=== Checking for deferred data in mail: {mail.Subject} ===");
                System.Diagnostics.Debug.WriteLine($"Total UserProperties count: {props.Count}");
                foreach (Outlook.UserProperty p in props)
                {
                    System.Diagnostics.Debug.WriteLine($"  Property: {p.Name} = {p.Value}");
                }

                var prop = props.Find(DEFERRED_PROPERTY_NAME);

                if (prop == null)
                {
                    System.Diagnostics.Debug.WriteLine($"No deferred property '{DEFERRED_PROPERTY_NAME}' found");
                    return null;
                }

                if (string.IsNullOrEmpty(prop.Value?.ToString()))
                {
                    System.Diagnostics.Debug.WriteLine($"Deferred property '{DEFERRED_PROPERTY_NAME}' exists but is empty");
                    return null;
                }

                var json = prop.Value.ToString();
                System.Diagnostics.Debug.WriteLine($"✓ Found deferred property: {json}");

                var data = JsonSerializer.Deserialize<DeferredActionData>(json);
                System.Diagnostics.Debug.WriteLine($"✓ Deserialized: Mode={data.Mode}, ActionID={data.ActionID}");

                return data;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR loading deferred data: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
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
                    System.Diagnostics.Debug.WriteLine("Cleared deferred data from sent item");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing deferred data: {ex.Message}");
            }
        }

        /// <summary>
        /// Executes a deferred action based on the deferred data
        /// </summary>
        private async Task ExecuteDeferredActionAsync(Outlook.MailItem mail, DeferredActionData data)
        {
            System.Diagnostics.Debug.WriteLine($"=== ExecuteDeferredActionAsync START ===");
            System.Diagnostics.Debug.WriteLine($"Mode: {data.Mode}, ActionID: {data.ActionID}");
            System.Diagnostics.Debug.WriteLine($"Mail subject: {mail.Subject}");

            try
            {
                if (_actionPane == null)
                {
                    System.Diagnostics.Debug.WriteLine("ERROR: ActionPane is null, cannot execute deferred action");
                    return;
                }

                // Execute based on mode
                if (data.Mode == "Create")
                {
                    System.Diagnostics.Debug.WriteLine("→ Dispatching ExecuteDeferredCreateAsync to UI thread");
                    await _actionPane.Dispatcher.InvokeAsync(async () =>
                    {
                        await _actionPane.ExecuteDeferredCreateAsync(mail);
                    });
                    System.Diagnostics.Debug.WriteLine("✓ ExecuteDeferredCreateAsync completed");
                }
                else if (data.Mode == "CreateMultiple")
                {
                    System.Diagnostics.Debug.WriteLine("→ Dispatching ExecuteDeferredCreateMultipleAsync to UI thread");
                    await _actionPane.Dispatcher.InvokeAsync(async () =>
                    {
                        await _actionPane.ExecuteDeferredCreateMultipleAsync(mail);
                    });
                    System.Diagnostics.Debug.WriteLine("✓ ExecuteDeferredCreateMultipleAsync completed");
                }
                else if (data.Mode == "Update" && data.ActionID.HasValue)
                {
                    System.Diagnostics.Debug.WriteLine($"→ Dispatching ExecuteDeferredUpdateAsync for action {data.ActionID.Value}");
                    await _actionPane.Dispatcher.InvokeAsync(async () =>
                    {
                        await _actionPane.ExecuteDeferredUpdateAsync(mail, data.ActionID.Value);
                    });
                    System.Diagnostics.Debug.WriteLine($"✓ ExecuteDeferredUpdateAsync completed for action {data.ActionID.Value}");
                }
                else if (data.Mode == "Close" && data.ActionID.HasValue)
                {
                    System.Diagnostics.Debug.WriteLine($"→ Dispatching ExecuteDeferredCloseAsync for action {data.ActionID.Value}");
                    await _actionPane.Dispatcher.InvokeAsync(async () =>
                    {
                        await _actionPane.ExecuteDeferredCloseAsync(mail, data.ActionID.Value);
                    });
                    System.Diagnostics.Debug.WriteLine($"✓ ExecuteDeferredCloseAsync completed for action {data.ActionID.Value}");
                }
                else if (data.Mode == "Reopen" && data.ActionID.HasValue)
                {
                    System.Diagnostics.Debug.WriteLine($"→ Dispatching ExecuteDeferredReopenAsync for action {data.ActionID.Value}");
                    await _actionPane.Dispatcher.InvokeAsync(async () =>
                    {
                        await _actionPane.ExecuteDeferredReopenAsync(mail, data.ActionID.Value);
                    });
                    System.Diagnostics.Debug.WriteLine($"✓ ExecuteDeferredReopenAsync completed for action {data.ActionID.Value}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"WARNING: Unknown mode '{data.Mode}' or missing ActionID");
                }

                // Cleanup: Remove the deferred property
                System.Diagnostics.Debug.WriteLine("→ Clearing deferred data property");
                ClearDeferredData(mail);
                System.Diagnostics.Debug.WriteLine("✓ Deferred data cleared");

                System.Diagnostics.Debug.WriteLine($"=== ExecuteDeferredActionAsync END ===");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR executing deferred action: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                // Note: We don't clear the deferred data on error, so user can retry manually
            }
        }

        #endregion

        #region InspectorWrapper Management

        /// <summary>
        /// Generates a unique key for an Inspector instance
        /// </summary>
        private string GetInspectorKey(Outlook.Inspector inspector)
        {
            return inspector.GetHashCode().ToString();
        }

        /// <summary>
        /// Retrieves the InspectorWrapper for a given Inspector
        /// </summary>
        public InspectorWrapper GetInspectorWrapper(Outlook.Inspector inspector)
        {
            string key = GetInspectorKey(inspector);
            return _inspectorWrappers.ContainsKey(key) ? _inspectorWrappers[key] : null;
        }

        /// <summary>
        /// Gets the action pane instance for ribbon access
        /// </summary>
        public UI.ProjectActionPane GetActionPane()
        {
            return _actionPane;
        }

        /// <summary>
        /// Called by InspectorWrapper when an Inspector closes
        /// </summary>
        public void OnInspectorClose(Outlook.Inspector inspector)
        {
            try
            {
                string key = GetInspectorKey(inspector);
                if (_inspectorWrappers.ContainsKey(key))
                {
                    _inspectorWrappers.Remove(key);
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper removed from dictionary: {key}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnInspectorClose: {ex.Message}");
            }
        }

        #endregion

        #region Ribbon Factory

        /// <summary>
        /// Creates the appropriate ribbon for the given context (Inspector or Explorer)
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            System.Diagnostics.Debug.WriteLine($"=== CreateRibbonExtensibilityObject CALLED at {DateTime.Now:HH:mm:ss.fff} ===");
            var ribbon = new InspectorComposeRibbon();
            System.Diagnostics.Debug.WriteLine($"    Created new ribbon instance: {ribbon.InstanceID}");
            return ribbon;
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}