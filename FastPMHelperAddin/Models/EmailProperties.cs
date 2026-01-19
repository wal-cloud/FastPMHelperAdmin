using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin.Models
{
    /// <summary>
    /// Contains extracted properties from an Outlook MailItem.
    /// Extracting these properties once on the COM thread avoids repeated synchronous calls
    /// on the UI thread, improving performance.
    /// </summary>
    public class EmailProperties
    {
        public string Subject { get; set; }
        public string InternetMessageId { get; set; }
        public string InReplyToId { get; set; }
        public string ConversationId { get; set; }
        public string StoreId { get; set; }
        public string EntryId { get; set; }
        public string Body { get; set; }
        public string SenderEmailAddress { get; set; }
        public string To { get; set; }
        public DateTime ReceivedTime { get; set; }
        public bool Sent { get; set; }

        // Keep reference to actual MailItem for operations that still need it
        public Outlook.MailItem MailItem { get; set; }

        /// <summary>
        /// Extracts all relevant properties from a MailItem on the current thread.
        /// This should be called on the COM thread to avoid cross-thread marshaling overhead.
        /// </summary>
        public static EmailProperties ExtractFrom(Outlook.MailItem mail)
        {
            if (mail == null)
                return null;

            var properties = new EmailProperties
            {
                MailItem = mail,
                Subject = mail.Subject ?? "",
                SenderEmailAddress = mail.SenderEmailAddress ?? "",
                To = mail.To ?? "",
                Sent = mail.Sent
            };

            // Extract ConversationID - may need special handling for drafts
            try
            {
                properties.ConversationId = mail.ConversationID;
                System.Diagnostics.Debug.WriteLine($"EmailProperties.ExtractFrom: ConversationID = '{properties.ConversationId ?? "(null)"}'");

                // For drafts, ConversationID might be null until saved
                // If null, try to get it by saving the draft first (only for unsent emails)
                if (string.IsNullOrEmpty(properties.ConversationId) && !mail.Sent)
                {
                    System.Diagnostics.Debug.WriteLine($"EmailProperties.ExtractFrom: ConversationID is null for draft, attempting save...");
                    try
                    {
                        mail.Save();
                        properties.ConversationId = mail.ConversationID;
                        System.Diagnostics.Debug.WriteLine($"EmailProperties.ExtractFrom: After save, ConversationID = '{properties.ConversationId ?? "(null)"}'");
                    }
                    catch (Exception saveEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"EmailProperties.ExtractFrom: Save failed: {saveEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"EmailProperties.ExtractFrom: Error getting ConversationID: {ex.Message}");
                properties.ConversationId = null;
            }

            try
            {
                properties.ReceivedTime = mail.ReceivedTime;
            }
            catch
            {
                properties.ReceivedTime = DateTime.MinValue;
            }

            // Extract Body (can be large, but necessary for classification)
            try
            {
                properties.Body = mail.Body ?? "";
            }
            catch
            {
                properties.Body = "";
            }

            // Extract InternetMessageId using PropertyAccessor
            try
            {
                const string PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
                var accessor = mail.PropertyAccessor;
                properties.InternetMessageId = accessor.GetProperty(PR_INTERNET_MESSAGE_ID)?.ToString() ?? "";
            }
            catch
            {
                properties.InternetMessageId = "";
            }

            // Extract InReplyTo using PropertyAccessor
            try
            {
                const string PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
                var accessor = mail.PropertyAccessor;
                properties.InReplyToId = accessor.GetProperty(PR_IN_REPLY_TO_ID)?.ToString() ?? "";
            }
            catch
            {
                properties.InReplyToId = "";
            }

            // Extract StoreID and EntryID for later retrieval
            try
            {
                properties.StoreId = mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : "";
            }
            catch
            {
                properties.StoreId = "";
            }

            try
            {
                properties.EntryId = mail.EntryID ?? "";
            }
            catch
            {
                properties.EntryId = "";
            }

            return properties;
        }
    }
}
