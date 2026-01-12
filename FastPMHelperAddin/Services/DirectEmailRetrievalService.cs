using System;
using System.Collections.Generic;
using System.Linq;
using FastPMHelperAddin.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin.Services
{
    public class DirectEmailRetrievalService
    {
        private Outlook.Application _app;
        private Outlook.NameSpace _ns;

        public DirectEmailRetrievalService()
        {
            _app = Globals.ThisAddIn.Application;
            _ns = _app.GetNamespace("MAPI");
        }

        /// <summary>
        /// Retrieve emails using direct references (StoreID|EntryID|InternetMessageId)
        /// </summary>
        public List<RelatedEmailItem> RetrieveEmailsByReferences(List<string> emailReferences)
        {
            var results = new List<RelatedEmailItem>();

            if (emailReferences == null || emailReferences.Count == 0)
                return results;

            foreach (var reference in emailReferences)
            {
                // Parse reference format: StoreID|EntryID|InternetMessageId
                var parts = reference.Split(new[] { '|' }, StringSplitOptions.None);

                // Ignore legacy format (no pipes)
                if (parts.Length != 3)
                {
                    System.Diagnostics.Debug.WriteLine($"Ignoring legacy reference format: {reference}");
                    continue;
                }

                string storeId = parts[0].Trim();
                string entryId = parts[1].Trim();
                string internetMessageId = parts[2].Trim();

                // Validate components
                if (string.IsNullOrWhiteSpace(storeId) ||
                    string.IsNullOrWhiteSpace(entryId) ||
                    string.IsNullOrWhiteSpace(internetMessageId))
                {
                    System.Diagnostics.Debug.WriteLine($"Invalid reference components: {reference}");
                    continue;
                }

                // Try to retrieve email
                var mail = RetrieveEmail(storeId, entryId, internetMessageId);
                if (mail != null)
                {
                    results.Add(CreateRelatedEmailItem(mail, internetMessageId));
                }
            }

            // Sort by ReceivedTime descending (newest first)
            return results.OrderByDescending(e => e.ReceivedTime).ToList();
        }

        /// <summary>
        /// Retrieve a single email using direct reference with fallback
        /// </summary>
        private Outlook.MailItem RetrieveEmail(string storeId, string entryId, string internetMessageId)
        {
            // PRIMARY: Try direct retrieval using EntryID and StoreID
            try
            {
                object item = _ns.GetItemFromID(entryId, storeId);
                if (item is Outlook.MailItem mail)
                {
                    System.Diagnostics.Debug.WriteLine($"Retrieved email via EntryID: {mail.Subject}");
                    return mail;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetItemFromID failed (EntryID stale?): {ex.Message}");
            }

            // FALLBACK: Search by InternetMessageId if direct retrieval failed
            System.Diagnostics.Debug.WriteLine($"Falling back to search for: {internetMessageId}");
            return SearchByInternetMessageId(internetMessageId);
        }

        /// <summary>
        /// Fallback search by InternetMessageId across Inbox and Sent Items
        /// </summary>
        private Outlook.MailItem SearchByInternetMessageId(string targetMessageId)
        {
            string normalizedTarget = NormalizeMessageId(targetMessageId);

            try
            {
                // Search across all stores
                foreach (Outlook.Store store in _ns.Stores)
                {
                    try
                    {
                        var rootFolder = store.GetRootFolder();

                        // Search Inbox
                        var mail = SearchFolder(GetInboxFolder(rootFolder), normalizedTarget);
                        if (mail != null)
                            return mail;

                        // Search Sent Items
                        mail = SearchFolder(GetSentItemsFolder(rootFolder), normalizedTarget);
                        if (mail != null)
                            return mail;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error searching store: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in fallback search: {ex.Message}");
            }

            return null;
        }

        private Outlook.MailItem SearchFolder(Outlook.MAPIFolder folder, string normalizedTargetId)
        {
            if (folder == null)
                return null;

            try
            {
                foreach (object item in folder.Items)
                {
                    if (item is Outlook.MailItem mail)
                    {
                        try
                        {
                            string messageId = GetInternetMessageId(mail);
                            if (!string.IsNullOrEmpty(messageId))
                            {
                                string normalized = NormalizeMessageId(messageId);
                                if (normalized == normalizedTargetId)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Found via search: {mail.Subject}");
                                    return mail;
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            return null;
        }

        private RelatedEmailItem CreateRelatedEmailItem(Outlook.MailItem mail, string internetMessageId)
        {
            // Detect if this is a sent email
            bool isSent = false;
            try
            {
                var folder = mail.Parent as Outlook.MAPIFolder;
                if (folder != null)
                {
                    isSent = folder.Name.Equals("Sent Items", StringComparison.OrdinalIgnoreCase) ||
                             folder.Name.Equals("Sent", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch { }

            // Extract recipient names for sent emails
            string toRecipients = "";
            if (isSent)
            {
                try
                {
                    var recipientNames = new System.Collections.Generic.List<string>();
                    foreach (Outlook.Recipient recipient in mail.Recipients)
                    {
                        string name = recipient.Name ?? recipient.Address ?? "";
                        if (!string.IsNullOrEmpty(name))
                            recipientNames.Add(name);
                    }
                    toRecipients = string.Join(", ", recipientNames);
                }
                catch { }
            }

            return new RelatedEmailItem
            {
                ReceivedTime = mail.ReceivedTime,
                SenderName = mail.SenderName ?? "",
                SenderEmail = mail.SenderEmailAddress ?? "",
                Subject = mail.Subject ?? "",
                InternetMessageId = internetMessageId,
                IsSent = isSent,
                ToRecipients = toRecipients,
                MailItem = mail
            };
        }

        private string GetInternetMessageId(Outlook.MailItem mail)
        {
            const string PR_INTERNET_MESSAGE_ID =
                "http://schemas.microsoft.com/mapi/proptag/0x1035001E";

            try
            {
                var accessor = mail.PropertyAccessor;
                object value = accessor.GetProperty(PR_INTERNET_MESSAGE_ID);
                return value?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string NormalizeMessageId(string messageId)
        {
            if (string.IsNullOrWhiteSpace(messageId))
                return string.Empty;

            messageId = messageId.Trim();
            if (messageId.StartsWith("<") && messageId.EndsWith(">"))
                messageId = messageId.Substring(1, messageId.Length - 2);

            return messageId;
        }

        private Outlook.MAPIFolder GetInboxFolder(Outlook.MAPIFolder rootFolder)
        {
            return FindFolderByName(rootFolder, "Inbox");
        }

        private Outlook.MAPIFolder GetSentItemsFolder(Outlook.MAPIFolder rootFolder)
        {
            return FindFolderByName(rootFolder, "Sent Items");
        }

        private Outlook.MAPIFolder FindFolderByName(Outlook.MAPIFolder parentFolder, string folderName)
        {
            try
            {
                foreach (Outlook.MAPIFolder folder in parentFolder.Folders)
                {
                    if (folder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                        return folder;
                }
            }
            catch { }
            return null;
        }
    }
}
