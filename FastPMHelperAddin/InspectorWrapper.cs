using System;
using System.Linq;
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

            // CRITICAL: Force ribbon to refresh with new grouping
            // This is needed because Outlook may reuse an existing ribbon instance
            // and won't call GetActionCount unless we invalidate
            InspectorComposeRibbon.InvalidateCurrentRibbon();
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
                var groupingService = actionPane.GroupingService;
                var dashboardAction = actionPane.DashboardSelectedAction;
                var sidebarEmailProps = actionPane.CurrentEmailProperties;

                if (openActions == null || groupingService == null)
                {
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper: Cannot capture grouping - openActions={openActions != null}, groupingService={groupingService != null}");
                    return;
                }

                // Extract properties from the actual mail item being composed
                Models.EmailProperties mailItemProps = null;
                if (_mailItem != null)
                {
                    mailItemProps = Models.EmailProperties.ExtractFrom(_mailItem);
                }

                // Check for stored email-to-action mapping (highest priority)
                // This preserves the action selected when the email was opened
                Models.ActionItem storedAction = null;
                if (mailItemProps != null && !string.IsNullOrEmpty(mailItemProps.InReplyToId))
                {
                    storedAction = Globals.ThisAddIn.GetAndClearEmailActionMapping(mailItemProps.InReplyToId);
                    if (storedAction != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"InspectorWrapper: Found stored action mapping for this reply - using Action {storedAction.Id} ({storedAction.Title})");
                        dashboardAction = storedAction; // Override current dashboard action with stored one
                    }
                }

                // Determine the workflow type and appropriate context
                string packageContext = "";
                string projectContext = "";
                string internetMessageId = "";
                string inReplyToId = "";
                string conversationId = "";

                if (dashboardAction != null)
                {
                    // WORKFLOW 1: Open Actions double-click → Reply
                    // Use dashboard action context + mail item properties
                    packageContext = dashboardAction.Package ?? "";
                    projectContext = dashboardAction.Project ?? "";

                    if (mailItemProps != null)
                    {
                        internetMessageId = mailItemProps.InternetMessageId;
                        inReplyToId = mailItemProps.InReplyToId;
                        conversationId = mailItemProps.ConversationId;
                    }

                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper: WORKFLOW 1 - Open Actions reply");
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper:   Package: '{packageContext}', Project: '{projectContext}'");
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper:   InReplyTo: '{inReplyToId}'");
                }
                else if (sidebarEmailProps != null && mailItemProps != null &&
                         IsMatchingSidebarEmail(sidebarEmailProps, mailItemProps))
                {
                    // WORKFLOW 2: Popout from sidebar
                    // Sidebar email matches this mail item - use sidebar context and properties
                    packageContext = actionPane.CurrentPackageContext ?? "";
                    projectContext = actionPane.CurrentProjectContext ?? "";
                    internetMessageId = sidebarEmailProps.InternetMessageId;
                    inReplyToId = sidebarEmailProps.InReplyToId;
                    conversationId = sidebarEmailProps.ConversationId;

                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper: WORKFLOW 2 - Popout from sidebar");
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper:   Package: '{packageContext}', Project: '{projectContext}'");
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper:   Using sidebar email properties");
                }
                else
                {
                    // WORKFLOW 3: Reply from Inbox/Sent (or sidebar doesn't match)
                    // Use mail item properties for everything
                    if (mailItemProps != null)
                    {
                        internetMessageId = mailItemProps.InternetMessageId;
                        inReplyToId = mailItemProps.InReplyToId;
                        conversationId = mailItemProps.ConversationId;
                    }

                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper: WORKFLOW 3 - Reply from Inbox/Sent");
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper:   InReplyTo: '{inReplyToId}'");
                    System.Diagnostics.Debug.WriteLine($"InspectorWrapper:   No package/project context (will be detected from linked actions)");
                }

                // Calculate grouping
                CapturedGrouping = groupingService.GroupActions(
                    openActions,
                    internetMessageId,
                    inReplyToId,
                    conversationId,
                    packageContext,
                    projectContext
                );

                // CRITICAL FIX for WORKFLOW 1: Ensure dashboard action is first in linked actions
                if (dashboardAction != null && CapturedGrouping.LinkedActions.Count > 1)
                {
                    // Check if dashboard action is in the linked actions list
                    var dashboardActionInList = CapturedGrouping.LinkedActions.FirstOrDefault(a => a.Id == dashboardAction.Id);
                    if (dashboardActionInList != null)
                    {
                        // Remove it from current position and insert at beginning
                        CapturedGrouping.LinkedActions.Remove(dashboardActionInList);
                        CapturedGrouping.LinkedActions.Insert(0, dashboardActionInList);
                        System.Diagnostics.Debug.WriteLine($"InspectorWrapper: Reordered linked actions - dashboard action {dashboardAction.Id} moved to first position");
                    }
                }

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

        /// <summary>
        /// Checks if the sidebar's email matches the mail item being composed.
        /// Used to detect if this is a popout from sidebar vs a reply from inbox/sent.
        /// </summary>
        private bool IsMatchingSidebarEmail(Models.EmailProperties sidebarProps, Models.EmailProperties mailItemProps)
        {
            // For compose mode (reply), the mail item's InReplyTo should match the sidebar's MessageID
            // Or the ConversationIDs should match
            if (!string.IsNullOrEmpty(mailItemProps.InReplyToId) &&
                !string.IsNullOrEmpty(sidebarProps.InternetMessageId))
            {
                string normalizedInReplyTo = NormalizeMessageId(mailItemProps.InReplyToId);
                string normalizedSidebarMsgId = NormalizeMessageId(sidebarProps.InternetMessageId);

                if (normalizedInReplyTo == normalizedSidebarMsgId)
                {
                    return true;
                }
            }

            // Fallback: Check ConversationID match
            if (!string.IsNullOrEmpty(mailItemProps.ConversationId) &&
                !string.IsNullOrEmpty(sidebarProps.ConversationId) &&
                mailItemProps.ConversationId == sidebarProps.ConversationId)
            {
                return true;
            }

            return false;
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
