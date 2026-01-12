using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms.Integration;
using FastPMHelperAddin.UI;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin
{
    public partial class ThisAddIn
    {
        private Outlook.Application _app;
        private Outlook.NameSpace _ns;
        private List<Outlook.Items> _monitoredInboxes = new List<Outlook.Items>();

        private MailExporter _exporter;
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private ProjectActionPane _actionPane;
        private Outlook.Explorer _currentExplorer;

        private const string RootPath = @"C:\MailPipeline";
        private const string TrackedEmailAccount = "wally.cloud@dynonobel.com";
        
        // StoreId identifiers for filtering (hex encoded email addresses in StoreId)
        private const string PersonalAccountStoreIdHex = "77616C6C79636C6F7564406F75746C6F6F6B2E636F6D"; // wallycloud@outlook.com
        private const string DynoAccountStoreIdHex = "57616C6C792E436C6F75644064796E6F6E6F62656C2E636F6D"; // Wally.Cloud@dynonobel.com

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                _app = Globals.ThisAddIn.Application;
                _ns = _app.GetNamespace("MAPI");

                // Initialize exporter
                _exporter = new MailExporter(RootPath, GetSenderEmailAddress);

                // Monitor ALL inboxes across all accounts
                System.Diagnostics.Debug.WriteLine("=== Setting up inbox monitoring ===");
                foreach (Outlook.Store store in _ns.Stores)
                {
                    try
                    {
                        var rootFolder = store.GetRootFolder();
                        var inbox = GetInboxFolder(rootFolder);

                        if (inbox != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"Monitoring inbox: {inbox.Name} in store: {store.DisplayName}");
                            var items = inbox.Items;
                            items.ItemAdd += InboxItems_ItemAdd;
                            _monitoredInboxes.Add(items);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error setting up monitoring for store {store.DisplayName}: {ex.Message}");
                    }
                }

                // Monitor Sent Items for ALL accounts
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

                        // Marshal to WPF UI thread (COM -> WPF threading)
                        _actionPane.Dispatcher.Invoke(() =>
                        {
                            _actionPane.OnEmailSelected(mail);
                        });
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("  Selection is not a MailItem");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  No selection - clearing pane");
                    // No selection - clear the pane
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

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                // Unhook events
                foreach (var items in _monitoredInboxes)
                {
                    if (items != null)
                        items.ItemAdd -= InboxItems_ItemAdd;
                }

                // NEW: Unhook explorer event
                if (_currentExplorer != null)
                {
                    _currentExplorer.SelectionChange -= Explorer_SelectionChange;
                    System.Diagnostics.Debug.WriteLine("Explorer SelectionChange event unhooked");
                }

                System.Diagnostics.Debug.WriteLine("MailPipeline add-in shut down successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during shutdown: {ex.Message}");
            }
        }

        private void InboxItems_ItemAdd(object item)
        {
            if (item is Outlook.MailItem mail)
            {
                try
                {
                    // Only track if received in the tracked account
                    if (IsToTrackedAccount(mail))
                    {
                        _exporter.ExportMail(mail, "Inbox");
                        System.Diagnostics.Debug.WriteLine($"Exported inbox mail: {mail.Subject}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error exporting inbox mail: {ex.Message}");
                }
            }
        }

        private void SentItems_ItemAdd(object item)
        {
            System.Diagnostics.Debug.WriteLine("=== SentItems_ItemAdd EVENT FIRED ===");

            if (item is Outlook.MailItem mail)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"Item is MailItem with subject: {mail.Subject}");

                    // Only track if sent from the tracked account
                    if (IsFromTrackedAccount(mail))
                    {
                        _exporter.ExportMail(mail, "Sent");
                        System.Diagnostics.Debug.WriteLine($"Exported sent mail: {mail.Subject}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Mail NOT from tracked account - skipping export");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error exporting sent mail: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Item is not a MailItem, type: {item?.GetType().Name ?? "null"}");
            }
        }

        private bool IsFromTrackedAccount(Outlook.MailItem mail)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Checking if mail is from tracked account: {mail.Subject}");

                // Method 1: Check StoreId to identify which account sent the email
                var folder = mail.Parent as Outlook.MAPIFolder;
                if (folder != null && !string.IsNullOrEmpty(folder.StoreID))
                {
                    System.Diagnostics.Debug.WriteLine($"  Mail StoreId: {folder.StoreID}");

                    // Check if this email is in the dyno account store
                    if (folder.StoreID.IndexOf(DynoAccountStoreIdHex, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        System.Diagnostics.Debug.WriteLine("  Found in dyno account store - tracking!");
                        return true;
                    }

                    // Explicitly reject personal account emails
                    if (folder.StoreID.IndexOf(PersonalAccountStoreIdHex, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        System.Diagnostics.Debug.WriteLine("  Found in personal account store - NOT tracking");
                        return false;
                    }
                }

                // Method 2: Fallback to sender address checking if StoreId method fails
                System.Diagnostics.Debug.WriteLine($"  StoreId check inconclusive, checking sender address...");
                string senderAddress = GetSenderEmailAddress(mail);
                System.Diagnostics.Debug.WriteLine($"  Sender address: {senderAddress}");

                bool matches = senderAddress.Equals(TrackedEmailAccount, StringComparison.OrdinalIgnoreCase);
                System.Diagnostics.Debug.WriteLine($"  Match result: {matches}");

                return matches;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking sender account: {ex.Message}");
                return false;
            }
        }

        private bool IsToTrackedAccount(Outlook.MailItem mail)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Checking if mail is to tracked account: {mail.Subject}");

                // Method 1: Check StoreId to identify which account received the email
                var folder = mail.Parent as Outlook.MAPIFolder;
                if (folder != null && !string.IsNullOrEmpty(folder.StoreID))
                {
                    System.Diagnostics.Debug.WriteLine($"  Mail StoreId: {folder.StoreID}");
                    
                    // Check if this email is in the dyno account store
                    if (folder.StoreID.IndexOf(DynoAccountStoreIdHex, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        System.Diagnostics.Debug.WriteLine("  Found in dyno account store - tracking!");
                        return true;
                    }
                    
                    // Explicitly reject personal account emails
                    if (folder.StoreID.IndexOf(PersonalAccountStoreIdHex, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        System.Diagnostics.Debug.WriteLine("  Found in personal account store - NOT tracking");
                        return false;
                    }
                }

                // Method 2: Fallback to recipient checking if StoreId method fails
                System.Diagnostics.Debug.WriteLine($"  StoreId check inconclusive, checking recipients...");
                System.Diagnostics.Debug.WriteLine($"  To: {mail.To}");
                System.Diagnostics.Debug.WriteLine($"  CC: {mail.CC}");

                // Check To field
                if (!string.IsNullOrEmpty(mail.To) &&
                    mail.To.IndexOf(TrackedEmailAccount, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    System.Diagnostics.Debug.WriteLine("  Found in To field!");
                    return true;
                }

                // Check CC field
                if (!string.IsNullOrEmpty(mail.CC) &&
                    mail.CC.IndexOf(TrackedEmailAccount, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    System.Diagnostics.Debug.WriteLine("  Found in CC field!");
                    return true;
                }

                // Method 3: Check each recipient object
                foreach (Outlook.Recipient recipient in mail.Recipients)
                {
                    string recipientAddress = GetRecipientEmailAddress(recipient);
                    System.Diagnostics.Debug.WriteLine($"  Checking recipient: {recipientAddress}");

                    if (recipientAddress.Equals(TrackedEmailAccount, StringComparison.OrdinalIgnoreCase))
                    {
                        System.Diagnostics.Debug.WriteLine("  Match found in recipients!");
                        return true;
                    }
                }

                System.Diagnostics.Debug.WriteLine("  No match found - not tracking this email");
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking recipient account: {ex.Message}");
                return false;
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