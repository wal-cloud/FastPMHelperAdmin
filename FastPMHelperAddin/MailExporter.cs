using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

public class MailExporter
{
    private readonly string _rootPath;
    private readonly Func<Outlook.MailItem, string> _getSenderEmailAddress;

    public MailExporter(string rootPath, Func<Outlook.MailItem, string> getSenderEmailAddress)
    {
        _rootPath = rootPath;
        _getSenderEmailAddress = getSenderEmailAddress;
    }

    public void ExportMail(Outlook.MailItem mail, string direction)
    {
        try
        {
            // 1) Build unique folder path
            string timestamp = DateTime.UtcNow.ToString("yyyyMMddTHHmmssfffZ");
            string shortId = Guid.NewGuid().ToString("N").Substring(0, 6).ToUpper();

            string baseDir = Path.Combine(
                _rootPath,
                direction,          // "Inbox" or "Sent"
                "Queue",
                $"{timestamp}_{shortId}"
            );

            System.Diagnostics.Debug.WriteLine($"Creating directory: {baseDir}");
            Directory.CreateDirectory(baseDir);

            // 2) Save attachments
            System.Diagnostics.Debug.WriteLine($"Saving {mail.Attachments.Count} attachments...");
            var attachmentFiles = SaveAttachments(mail, baseDir);

            // 3) Build metadata DTO
            System.Diagnostics.Debug.WriteLine("Building metadata...");
            var meta = BuildMeta(mail, direction, attachmentFiles, timestamp, shortId);

            // 4) Write meta.json
            System.Diagnostics.Debug.WriteLine("Writing meta.json...");
            string metaJson = JsonSerializer.Serialize(
                meta,
                new JsonSerializerOptions { WriteIndented = true });

            string metaPath = Path.Combine(baseDir, "meta.json");
            File.WriteAllText(metaPath, metaJson);

            System.Diagnostics.Debug.WriteLine($"Successfully exported to: {metaPath}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"ERROR in ExportMail: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }

    private List<string> SaveAttachments(Outlook.MailItem mail, string baseDir)
    {
        var list = new List<string>();
        int regularIdx = 1;
        int embeddedIdx = 1;

        // Create embedded subfolder if needed
        string embeddedDir = Path.Combine(baseDir, "embedded");

        System.Diagnostics.Debug.WriteLine($"=== Processing {mail.Attachments.Count} attachments ===");

        // Get HTML body once for all attachments to check references
        string htmlBody = string.Empty;
        try
        {
            if (mail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
            {
                htmlBody = mail.HTMLBody ?? string.Empty;
                System.Diagnostics.Debug.WriteLine($"Email is HTML format, body length: {htmlBody.Length}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Email format: {mail.BodyFormat} (not HTML)");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error getting HTML body: {ex.Message}");
        }

        foreach (Outlook.Attachment att in mail.Attachments)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Processing attachment: {att.FileName}");
                System.Diagnostics.Debug.WriteLine($"  Type: {att.Type}");
                System.Diagnostics.Debug.WriteLine($"  DisplayName: {att.DisplayName}");

                bool isEmbedded = IsEmbeddedAttachment(att, htmlBody);
                string safeName = MakeSafeFileName(att.FileName);

                string relativePath;
                string fullPath;

                if (isEmbedded)
                {
                    System.Diagnostics.Debug.WriteLine($"  >> EMBEDDED attachment detected");
                    // Create embedded directory if it doesn't exist
                    if (!Directory.Exists(embeddedDir))
                        Directory.CreateDirectory(embeddedDir);

                    string numberedName = $"{embeddedIdx:D2}_{safeName}";
                    relativePath = Path.Combine("embedded", numberedName);
                    fullPath = Path.Combine(embeddedDir, numberedName);
                    embeddedIdx++;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"  >> REGULAR attachment detected");
                    string numberedName = $"{regularIdx:D2}_{safeName}";
                    relativePath = numberedName;
                    fullPath = Path.Combine(baseDir, numberedName);
                    regularIdx++;
                }

                att.SaveAsFile(fullPath);
                list.Add(relativePath);
                System.Diagnostics.Debug.WriteLine($"  Saved to: {relativePath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving attachment {att.FileName}: {ex.Message}");
                // Continue with other attachments
            }
        }

        return list;
    }

    private bool IsEmbeddedAttachment(Outlook.Attachment attachment, string htmlBody)
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("    Checking if embedded...");

            string contentId = null;

            // Method 1: Check PR_ATTACH_CONTENT_ID and verify it's referenced in HTML body
            const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E";
            try
            {
                var contentIdObj = attachment.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID);
                if (contentIdObj != null && !string.IsNullOrEmpty(contentIdObj.ToString()))
                {
                    contentId = contentIdObj.ToString();
                    System.Diagnostics.Debug.WriteLine($"    PR_ATTACH_CONTENT_ID: {contentId}");

                    // If we have HTML body, check if this content ID is actually referenced
                    if (!string.IsNullOrEmpty(htmlBody))
                    {
                        // Check for cid: reference in the HTML (e.g., src="cid:image001.gif@01DC6528.308322F0")
                        bool isReferencedInHtml = htmlBody.IndexOf($"cid:{contentId}", StringComparison.OrdinalIgnoreCase) >= 0;

                        System.Diagnostics.Debug.WriteLine($"    Referenced in HTML body: {isReferencedInHtml}");

                        if (isReferencedInHtml)
                        {
                            System.Diagnostics.Debug.WriteLine($"    Content ID found in HTML -> EMBEDDED");
                            return true;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"    Content ID NOT found in HTML -> checking other methods");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"    No HTML body to check against");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"    PR_ATTACH_CONTENT_ID: (empty or null)");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"    PR_ATTACH_CONTENT_ID: Error - {ex.Message}");
            }

            // Method 2: Check PR_ATTACH_FLAGS for inline flag (4 = ATT_MHTML_REF)
            const string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";
            try
            {
                var flags = attachment.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS);
                if (flags != null)
                {
                    int flagValue = Convert.ToInt32(flags);
                    System.Diagnostics.Debug.WriteLine($"    PR_ATTACH_FLAGS: {flagValue} (0x{flagValue:X})");
                    if ((flagValue & 0x4) != 0) // Check if bit 2 is set (ATT_MHTML_REF)
                    {
                        System.Diagnostics.Debug.WriteLine($"    PR_ATTACH_FLAGS has ATT_MHTML_REF bit -> EMBEDDED");
                        return true;
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"    PR_ATTACH_FLAGS: (null)");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"    PR_ATTACH_FLAGS: Error - {ex.Message}");
            }

            // Method 3: Check PR_ATTACHMENT_HIDDEN
            const string PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";
            try
            {
                var hidden = attachment.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN);
                System.Diagnostics.Debug.WriteLine($"    PR_ATTACHMENT_HIDDEN: {hidden}");
                if (hidden != null && Convert.ToBoolean(hidden))
                {
                    System.Diagnostics.Debug.WriteLine($"    PR_ATTACHMENT_HIDDEN is true -> EMBEDDED");
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"    PR_ATTACHMENT_HIDDEN: Error - {ex.Message}");
            }

            System.Diagnostics.Debug.WriteLine($"    No embedded indicators found -> REGULAR");
            return false;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"    Error checking if attachment is embedded: {ex.Message}");
            return false; // Default to regular attachment if we can't determine
        }
    }

    private object BuildMeta(
        Outlook.MailItem mail,
        string direction,
        List<string> attachments,
        string timestamp,
        string shortId)
    {
        var folder = mail.Parent as Outlook.MAPIFolder;
        string storeId = folder?.StoreID ?? string.Empty;

        // Split body into current message and conversation history
        var bodyParts = SplitEmailBody(mail.Body ?? string.Empty);

        string referencesRaw = GetReferencesRaw(mail);

        return new
        {
            Id = $"{direction}/{timestamp}/{shortId}",
            Direction = direction,                 // "Inbox" / "Sent"
            CreatedAtUtc = timestamp,
            EntryId = mail.EntryID,
            StoreId = storeId,
            InternetMessageId = GetInternetMessageId(mail),
            ConversationId = mail.ConversationID,
            InReplyToId = GetInReplyToId(mail),
            ReferencesRaw = referencesRaw,
            References = ParseReferences(referencesRaw),
            Subject = mail.Subject ?? string.Empty,
            From = _getSenderEmailAddress != null ? _getSenderEmailAddress(mail) : mail.SenderEmailAddress ?? string.Empty,
            To = mail.To ?? string.Empty,
            Cc = mail.CC ?? string.Empty,
            Bcc = mail.BCC ?? string.Empty,
            SentOn = mail.SentOn,
            ReceivedTime = mail.ReceivedTime,

            // Structured body content
            Body = new
            {
                Full = mail.Body ?? string.Empty,              // Complete email thread
                Current = bodyParts.CurrentMessage,             // Just the most recent message
                HasHistory = bodyParts.HasHistory,              // True if there are previous messages
                History = bodyParts.History                     // Previous messages in the thread
            },

            Attachments = attachments
        };
    }

    private (string CurrentMessage, bool HasHistory, string History) SplitEmailBody(string fullBody)
    {
        if (string.IsNullOrEmpty(fullBody))
            return (string.Empty, false, string.Empty);

        // Common reply separators in emails
        string[] replySeparators = new[]
        {
                "\r\n________________________________\r\n",  // Outlook horizontal line
                "\r\nFrom:",                                 // Standard "From:" line
                "\r\n-----Original Message-----",            // Outlook original message
                "\r\nOn ",                                   // "On [date], [person] wrote:"
                "\r\n>",                                     // Quoted text marker
                "\r\n\r\nFrom:",                             // Double line break before From
            };

        int earliestSeparatorIndex = -1;
        string foundSeparator = null;

        // Find the earliest separator
        foreach (var separator in replySeparators)
        {
            int index = fullBody.IndexOf(separator, StringComparison.OrdinalIgnoreCase);
            if (index > 0 && (earliestSeparatorIndex == -1 || index < earliestSeparatorIndex))
            {
                earliestSeparatorIndex = index;
                foundSeparator = separator;
            }
        }

        // If we found a separator, split the content
        if (earliestSeparatorIndex > 0)
        {
            string currentMessage = fullBody.Substring(0, earliestSeparatorIndex).Trim();
            string history = fullBody.Substring(earliestSeparatorIndex).Trim();

            return (currentMessage, true, history);
        }

        // No separator found - entire body is the current message
        return (fullBody.Trim(), false, string.Empty);
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

    private string GetInReplyToId(Outlook.MailItem mail)
    {
        const string PR_IN_REPLY_TO_ID =
            "http://schemas.microsoft.com/mapi/proptag/0x1042001E";

        try
        {
            var accessor = mail.PropertyAccessor;
            object value = accessor.GetProperty(PR_IN_REPLY_TO_ID);
            return value?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private string GetReferencesRaw(Outlook.MailItem mail)
    {
        const string PR_INTERNET_REFERENCES =
            "http://schemas.microsoft.com/mapi/proptag/0x1039001E";

        try
        {
            var accessor = mail.PropertyAccessor;
            object value = accessor.GetProperty(PR_INTERNET_REFERENCES);
            return value?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private string[] ParseReferences(string referencesRaw)
    {
        if (string.IsNullOrWhiteSpace(referencesRaw))
            return new string[0];

        try
        {
            // References header contains space-separated message IDs
            // Each ID is typically in angle brackets: <id1@domain> <id2@domain>
            var references = new List<string>();
            var parts = referencesRaw.Split(new[] { ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in parts)
            {
                string trimmed = part.Trim();

                // Remove angle brackets if present
                if (trimmed.StartsWith("<") && trimmed.EndsWith(">"))
                {
                    trimmed = trimmed.Substring(1, trimmed.Length - 2);
                }

                if (!string.IsNullOrWhiteSpace(trimmed))
                {
                    references.Add(trimmed);
                }
            }

            return references.ToArray();
        }
        catch
        {
            return new string[0];
        }
    }

    private string MakeSafeFileName(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            return "unnamed";

        foreach (char c in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(c, '_');
        }

        // Also replace some additional problematic characters
        fileName = fileName.Replace('<', '_')
                           .Replace('>', '_')
                           .Replace(':', '_')
                           .Replace('"', '_')
                           .Replace('/', '_')
                           .Replace('\\', '_')
                           .Replace('|', '_')
                           .Replace('?', '_')
                           .Replace('*', '_');

        return fileName;
    }
}
