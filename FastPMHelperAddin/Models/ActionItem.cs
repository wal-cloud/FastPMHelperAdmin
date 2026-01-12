using System;
using System.Collections.Generic;
using System.Linq;

namespace FastPMHelperAddin.Models
{
    public class ActionItem
    {
        // SharePoint List Item ID
        public int Id { get; set; }

        // Classification Properties
        public string Project { get; set; }    // Column B
        public string Package { get; set; }    // Column C

        // Core Properties
        public string Title { get; set; }      // Column D (moved from B)
        public string Status { get; set; }     // Column E (moved from C)
        public string BallHolder { get; set; } // Column F (moved from D)

        // Date Properties
        public DateTime? SentOn { get; set; }  // Column G (moved from E)
        public DateTime? DueDate { get; set; } // Column H (moved from F)

        public string HistoryLog { get; set; }      // Column I (moved from G)

        // Email Threading Properties (semicolon-separated)
        public string LinkedThreadIDs { get; set; }      // ConversationIDs
        public string ActiveMessageIDs { get; set; }     // InternetMessageIDs

        // Parsed Collections (cached)
        private List<string> _linkedThreadIds;
        private List<string> _activeMessageIds;

        public List<string> ParseLinkedThreadIds()
        {
            if (_linkedThreadIds == null)
                _linkedThreadIds = ParseIds(LinkedThreadIDs);
            return _linkedThreadIds;
        }

        public List<string> ParseActiveMessageIds()
        {
            if (_activeMessageIds == null)
                _activeMessageIds = ParseIds(ActiveMessageIDs);
            return _activeMessageIds;
        }

        public static List<string> ParseIds(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
                return new List<string>();

            return rawValue
                .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(id => id.Trim())
                .Where(id => !string.IsNullOrEmpty(id))
                .ToList();
        }

        // Display name for ComboBox
        public override string ToString()
        {
            return $"{Title} [{BallHolder}]";
        }
    }
}
