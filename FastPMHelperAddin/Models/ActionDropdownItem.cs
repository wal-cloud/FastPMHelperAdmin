using System;

namespace FastPMHelperAddin.Models
{
    public enum ActionDropdownItemType
    {
        Header,          // Section header - not selectable
        Action,          // Actual action item - selectable
        MoreExpander     // "More..." toggle - clickable but not selectable as value
    }

    public class ActionDropdownItem
    {
        public ActionDropdownItemType ItemType { get; set; }

        // For Header type
        public string HeaderText { get; set; }

        // For Action type
        public ActionItem Action { get; set; }

        // Category for filtering: "Linked", "Package", "Project", "Other"
        public string Category { get; set; }

        // Display text (combines action title with visual cues)
        public string DisplayText
        {
            get
            {
                if (ItemType == ActionDropdownItemType.Header)
                    return HeaderText;
                if (ItemType == ActionDropdownItemType.MoreExpander)
                    return "More...";
                return Action?.Title ?? "";
            }
        }

        // For ComboBox item equality
        public override bool Equals(object obj)
        {
            if (obj is ActionDropdownItem other && ItemType == ActionDropdownItemType.Action)
                return Action?.Id == other.Action?.Id;
            return false;
        }

        public override int GetHashCode()
        {
            return Action?.Id.GetHashCode() ?? 0;
        }
    }
}
