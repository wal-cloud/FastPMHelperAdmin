using System.Collections.Generic;
using System.Linq;
using FastPMHelperAddin.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin.Services
{
    public class ActionMatchingService
    {
        public ActionItem FindMatchingAction(Outlook.MailItem mail,
            List<ActionItem> openActions)
        {
            if (mail == null || openActions == null || openActions.Count == 0)
                return null;

            string inReplyToId = GetInReplyToId(mail);
            string conversationId = mail.ConversationID;

            // Strategy A: Hard Match - Check In-Reply-To against ActiveMessageIDs
            if (!string.IsNullOrEmpty(inReplyToId))
            {
                var hardMatch = openActions.FirstOrDefault(action =>
                {
                    var activeIds = action.ParseActiveMessageIds();
                    return activeIds.Any(id =>
                        NormalizeMessageId(id) == NormalizeMessageId(inReplyToId));
                });

                if (hardMatch != null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Hard match found: InReplyTo '{inReplyToId}' matches action '{hardMatch.Title}'");
                    return hardMatch;
                }
            }

            // Strategy B: Soft Match - Check ConversationID against LinkedThreadIDs
            if (!string.IsNullOrEmpty(conversationId))
            {
                var softMatch = openActions.FirstOrDefault(action =>
                {
                    var linkedThreads = action.ParseLinkedThreadIds();
                    return linkedThreads.Any(threadId =>
                        threadId == conversationId);
                });

                if (softMatch != null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Soft match found: ConversationID '{conversationId}' matches action '{softMatch.Title}'");
                    return softMatch;
                }
            }

            System.Diagnostics.Debug.WriteLine("No match found");
            return null;
        }

        private string GetInReplyToId(Outlook.MailItem mail)
        {
            const string PR_IN_REPLY_TO_ID =
                "http://schemas.microsoft.com/mapi/proptag/0x1042001E";

            try
            {
                var accessor = mail.PropertyAccessor;
                object value = accessor.GetProperty(PR_IN_REPLY_TO_ID);
                return NormalizeMessageId(value?.ToString() ?? string.Empty);
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

            // Remove angle brackets if present
            messageId = messageId.Trim();
            if (messageId.StartsWith("<") && messageId.EndsWith(">"))
                messageId = messageId.Substring(1, messageId.Length - 2);

            return messageId;
        }
    }
}
