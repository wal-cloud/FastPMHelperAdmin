using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin.Models
{
    public class RelatedEmailItem
    {
        public DateTime ReceivedTime { get; set; }
        public string SenderName { get; set; }
        public string SenderEmail { get; set; }
        public string Subject { get; set; }
        public string InternetMessageId { get; set; }
        public bool IsSent { get; set; }
        public string ToRecipients { get; set; }

        // Store reference to actual mail item for opening
        public Outlook.MailItem MailItem { get; set; }

        // Display properties - show TO: for sent emails, otherwise show sender name only
        public string FromDisplay
        {
            get
            {
                if (IsSent)
                {
                    return string.IsNullOrEmpty(ToRecipients)
                        ? "TO: (Unknown)"
                        : $"TO: {ToRecipients}";
                }
                else
                {
                    // Show only name, not email address
                    return string.IsNullOrEmpty(SenderName)
                        ? SenderEmail
                        : SenderName;
                }
            }
        }

        public string DateDisplay => ReceivedTime.ToString("dd/MM/yyyy HH:mm");

        // Keep RE:/FW: prefixes in subject
        public string SubjectDisplay
        {
            get
            {
                if (string.IsNullOrEmpty(Subject))
                    return "(No Subject)";

                return Subject.Trim();
            }
        }
    }
}
