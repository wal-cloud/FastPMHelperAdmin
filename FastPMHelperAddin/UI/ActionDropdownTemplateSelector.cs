using System.Windows;
using System.Windows.Controls;
using FastPMHelperAddin.Models;

namespace FastPMHelperAddin.UI
{
    public class ActionDropdownTemplateSelector : DataTemplateSelector
    {
        public DataTemplate HeaderTemplate { get; set; }
        public DataTemplate ActionTemplate { get; set; }
        public DataTemplate MoreExpanderTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (item is ActionDropdownItem dropdownItem)
            {
                switch (dropdownItem.ItemType)
                {
                    case ActionDropdownItemType.Header:
                        return HeaderTemplate;
                    case ActionDropdownItemType.MoreExpander:
                        return MoreExpanderTemplate;
                    default:
                        return ActionTemplate;
                }
            }
            return base.SelectTemplate(item, container);
        }
    }
}
