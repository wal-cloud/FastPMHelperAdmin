using System;
using System.Runtime.InteropServices;
using FastPMHelperAddin.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin
{
    /// <summary>
    /// Wraps an Inspector window for lifecycle management.
    /// Manages the Inspector's Close event and coordinates with the custom Ribbon.
    /// </summary>
    public class InspectorWrapper
    {
        private Outlook.Inspector _inspector;
        private Outlook.MailItem _mailItem;
        private InspectorComposeRibbon _ribbon;

        /// <summary>
        /// Captured grouping results from the sidebar at the time of pop-out.
        /// This ensures the ribbon shows the correct linked actions even if the
        /// user selects a different email in the Explorer after popping out.
        /// </summary>
        public ActionGroupingResult CapturedGrouping { get; private set; }

        public Outlook.Inspector Inspector => _inspector;
        public Outlook.MailItem MailItem => _mailItem;
        public InspectorComposeRibbon Ribbon
        {
            get => _ribbon;
            set => _ribbon = value;
        }

        public InspectorWrapper(Outlook.Inspector inspector)
        {
            _inspector = inspector;

            // Get the MailItem from the Inspector
            if (inspector.CurrentItem is Outlook.MailItem mailItem)
            {
                _mailItem = mailItem;
            }

            // Hook the Close event for cleanup
            ((Outlook.InspectorEvents_10_Event)_inspector).Close += InspectorWrapper_Close;

            // Capture the sidebar's current grouping results at the moment of pop-out
            CaptureGroupingFromSidebar();

            System.Diagnostics.Debug.WriteLine($"InspectorWrapper created for Inspector {inspector.GetHashCode()}");
        }

        /// <summary>
        /// Captures the sidebar's current grouping results so the ribbon can use them.
        /// This is called at pop-out time, before the user can change the Explorer selection.
        /// </summary>
        private void CaptureGroupingFromSidebar()
        {
            try
            {
                var actionPane = Globals.ThisAddIn.GetActionPane();
                if (actionPane == null)
                {
                    System.Diagnostics.Debug.WriteLine("InspectorWrapper: actionPane is null, cannot capture grouping");
                    return;
                }

                var openActions = actionPane.OpenActions;
                var emailProps = actionPane.CurrentEmailProperties;
                var groupingService = actionPane.GroupingService;

                if (openActions == null || emailProps == null || groupingService == null)
                {
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper: Cannot capture grouping - openActions={openActions != null}, emailProps={emailProps != null}, groupingService={groupingService != null}");
                    return;
                }

                // Calculate grouping using the sidebar's current email context
                CapturedGrouping = groupingService.GroupActions(
                    openActions,
                    emailProps.InternetMessageId,
                    emailProps.InReplyToId,
                    emailProps.ConversationId,
                    actionPane.CurrentPackageContext ?? "",
                    actionPane.CurrentProjectContext ?? ""
                );

                System.Diagnostics.Debug.WriteLine($"InspectorWrapper: Captured grouping - {CapturedGrouping.LinkedActions.Count} linked, {CapturedGrouping.PackageActions.Count} package, {CapturedGrouping.ProjectActions.Count} project, {CapturedGrouping.OtherActions.Count} other");

                if (CapturedGrouping.LinkedActions.Count > 0)
                {
                    foreach (var action in CapturedGrouping.LinkedActions)
                    {
                        System.Diagnostics.Debug.WriteLine($"InspectorWrapper:   Linked: ID {action.Id} - {action.Title}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"InspectorWrapper: Error capturing grouping: {ex.Message}");
            }
        }

        private void InspectorWrapper_Close()
        {
            System.Diagnostics.Debug.WriteLine($"InspectorWrapper closing for Inspector {_inspector?.GetHashCode()}");

            try
            {
                // Notify ThisAddIn to remove this wrapper from the dictionary
                Globals.ThisAddIn.OnInspectorClose(_inspector);

                // Unhook the Close event
                if (_inspector != null)
                {
                    ((Outlook.InspectorEvents_10_Event)_inspector).Close -= InspectorWrapper_Close;
                }

                // Release COM objects
                if (_mailItem != null)
                {
                    Marshal.ReleaseComObject(_mailItem);
                    _mailItem = null;
                }

                if (_inspector != null)
                {
                    Marshal.ReleaseComObject(_inspector);
                    _inspector = null;
                }

                _ribbon = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in InspectorWrapper_Close: {ex.Message}");
            }
        }
    }
}
