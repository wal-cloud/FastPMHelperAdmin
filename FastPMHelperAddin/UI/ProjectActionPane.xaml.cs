using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using FastPMHelperAddin.Models;
using FastPMHelperAddin.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastPMHelperAddin.UI
{
    public partial class ProjectActionPane : UserControl
    {
        private GoogleSheetsAuthService _googleAuthService;
        private GoogleSheetsService _googleSheetsService;
        private LLMService _llmService;
        private ActionMatchingService _matchingService;
        private AutoClassifierService _classifierService;
        private DirectEmailRetrievalService _emailRetrievalService;
        private ActionGroupingService _groupingService;
        private EmailCategoryService _emailCategoryService;

        private List<ActionItem> _openActions;
        private Outlook.MailItem _currentMail;
        private EmailProperties _currentEmailProperties; // Cached email properties to avoid redundant COM calls
        private List<ActionDropdownItem> _dropdownItems;
        private bool _isExpanded = false;
        private string _currentPackageContext;
        private string _currentProjectContext;

        // Public properties for ribbon access
        public List<ActionItem> OpenActions => _openActions;
        public ActionGroupingService GroupingService => _groupingService;
        public EmailProperties CurrentEmailProperties => _currentEmailProperties;
        public string CurrentPackageContext => _currentPackageContext;
        public string CurrentProjectContext => _currentProjectContext;

        // Queue for processing actions
        private Queue<Func<Task>> _actionQueue = new Queue<Func<Task>>();
        private bool _isProcessingQueue = false;

        // Change tracking for linked action fields
        private class ActionFieldState
        {
            public string Project { get; set; }
            public string Package { get; set; }
            public string Title { get; set; }
            public string BallHolder { get; set; }
            public DateTime? DueDate { get; set; }
        }

        private ActionFieldState _previousActionState;

        // User name for health indicator assignment checking
        private string _currentUserName;

        // Dashboard state
        private List<ActionItem> _overdueActions;
        private List<ActionItem> _withMeActions;
        private ActionItem _dashboardSelectedAction;

        // User filter state
        private bool _isUserFilterExpanded = false;
        private string _selectedFilterUser = null; // null = "With Me" mode, otherwise filtered user name

        // Collapse/Expand state
        private bool _isOpenActionsCollapsed = false;
        private bool _isLinkedActionCollapsed = false;

        // Size memory for collapse/expand
        private GridLength _savedLinkedActionHeight;
        private GridLength _savedOpenActionsHeight;

        // Compose mode fields for deferred action execution
        private bool _isComposeMode = false;
        private Outlook.MailItem _composeMail;
        private Outlook.Inspector _composeInspector;
        private DeferredActionData _currentDeferredData;
        private const string DEFERRED_PROPERTY_NAME = "FastPMDeferredAction"; // No underscores - Outlook doesn't allow them
        private static readonly SolidColorBrush ScheduledBrush = new SolidColorBrush(Color.FromRgb(76, 175, 80)); // Green #4CAF50

        // Public properties to expose compose mode state (needed by ThisAddIn for race condition handling)
        public bool IsComposeMode => _isComposeMode;
        public Outlook.MailItem ComposeMail => _composeMail;
        public Outlook.Inspector ComposeInspector => _composeInspector;

        public ProjectActionPane()
        {
            InitializeComponent();
            LoadUserSettings();
            InitializeServices();
            LoadActionsAsync();
        }

        private void LoadUserSettings()
        {
            try
            {
                // Load user name from settings
                _currentUserName = Properties.Settings.Default.UserName;

                // If setting was null/empty, set default
                if (string.IsNullOrWhiteSpace(_currentUserName))
                {
                    _currentUserName = "Wally Cloud";
                    Properties.Settings.Default.UserName = _currentUserName;
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading user settings: {ex.Message}");
                _currentUserName = "Wally Cloud"; // Fallback
            }
        }

        // MyNameInput is now in SettingsWindow, so handler removed from here

        private void InitializeServices()
        {
            try
            {
                var config = Configuration.ConfigurationManager.Instance;
                config.ValidateGoogleSheetsConfiguration();

                // Initialize Google Sheets authentication
                _googleAuthService = new GoogleSheetsAuthService(
                    config.GoogleClientId,
                    config.GoogleClientSecret,
                    config.GoogleAppName,
                    config.GoogleTokenCacheDir
                );

                // Initialize Google Sheets service
                _googleSheetsService = new GoogleSheetsService(
                    _googleAuthService,
                    config.GoogleSpreadsheetsId,
                    config.GoogleSheetName
                );

                _llmService = new LLMService(config.GeminiApiKey);
                _matchingService = new ActionMatchingService();
                _classifierService = new AutoClassifierService();
                _emailRetrievalService = new DirectEmailRetrievalService();
                _groupingService = new ActionGroupingService();
                _emailCategoryService = new EmailCategoryService();

                System.Diagnostics.Debug.WriteLine("Services initialized successfully with Google Sheets");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to initialize services: {ex.Message}");
                MessageBox.Show(
                    $"Configuration error: {ex.Message}\n\nPlease ensure .env file contains:\n" +
                    "- GOOGLE_CLIENT_ID\n" +
                    "- GOOGLE_CLIENT_SECRET\n" +
                    "- GOOGLE_SHEETS_SPREADSHEET_ID\n" +
                    "- GEMINI_API_KEY\n\n" +
                    "See README.md for setup instructions.",
                    "Configuration Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                throw;
            }
        }

        private async Task LoadActionsAsync()
        {
            try
            {
                // Load classification rules
                var rulesData = await _googleSheetsService.FetchConfigRulesAsync();
                _classifierService.LoadRules(rulesData);
                System.Diagnostics.Debug.WriteLine($"Loaded {_classifierService.GetRuleCount()} classification rules");

                // Load actions
                _openActions = await _googleSheetsService.FetchOpenActionsAsync();

                // Update UI on the UI thread
                Dispatcher.Invoke(() =>
                {
                    RefreshActionComboBox();
                    RefreshDashboard();
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Load actions error: {ex.Message}");
                SetStatus("Error loading actions from Google Sheets.");
            }
        }

        private void RefreshActionComboBox()
        {
            ActionComboBox.ItemsSource = null;
            ActionComboBox.ItemsSource = _openActions;
        }

        private void RefreshDashboard()
        {
            if (_openActions == null || _openActions.Count == 0)
            {
                _overdueActions = new List<ActionItem>();
                _withMeActions = new List<ActionItem>();
                OverdueActionsComboBox.ItemsSource = null;
                WithMeActionsComboBox.ItemsSource = null;
                OverdueSection.Visibility = Visibility.Collapsed;
                WithMeSection.Visibility = Visibility.Collapsed;
                return;
            }

            // Filter overdue actions
            _overdueActions = _openActions
                .Where(a =>
                {
                    string status = a.Status ?? "";
                    return !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                        && a.DueDate.HasValue
                        && a.DueDate.Value.Date < DateTime.Today;
                })
                .OrderBy(a => a.DueDate)
                .ToList();

            // Filter with me/user actions (exclude overdue to avoid duplicates)
            // If user filter is active, filter by selected user; otherwise filter by current user
            string targetUser = _isUserFilterExpanded && !string.IsNullOrWhiteSpace(_selectedFilterUser)
                ? _selectedFilterUser
                : _currentUserName;

            bool isFilterMode = _isUserFilterExpanded && !string.IsNullOrWhiteSpace(_selectedFilterUser);

            _withMeActions = _openActions
                .Where(a =>
                {
                    string status = a.Status ?? "";
                    bool isOverdue = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                        && a.DueDate.HasValue
                        && a.DueDate.Value.Date < DateTime.Today;

                    bool isWithUser = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrWhiteSpace(a.BallHolder)
                        && (isFilterMode
                            ? IsUserInBallHolder(a.BallHolder, targetUser)
                            : IsAssignedToUser(a.BallHolder, targetUser));

                    return isWithUser && !isOverdue;
                })
                .OrderBy(a => a.DueDate)
                .ToList();

            // Update labels with counts
            OverdueLabel.Text = $"Overdue: {_overdueActions.Count}";

            if (isFilterMode)
            {
                WithMeLabel.Text = $"With {GetShortName(targetUser)}: {_withMeActions.Count}";
            }
            else
            {
                WithMeLabel.Text = $"With Me: {_withMeActions.Count}";
            }

            // Bind to UI
            OverdueActionsComboBox.ItemsSource = _overdueActions;
            WithMeActionsComboBox.ItemsSource = _withMeActions;

            // Update visibility
            OverdueSection.Visibility = _overdueActions.Count > 0
                ? Visibility.Visible
                : Visibility.Collapsed;

            WithMeSection.Visibility = _withMeActions.Count > 0
                ? Visibility.Visible
                : Visibility.Collapsed;

            // Update collapsed summary if Open Actions is collapsed
            // Always show "With me:" and calculate count based on current user
            if (_isOpenActionsCollapsed)
            {
                int myActionsCount = _openActions
                    .Where(a =>
                    {
                        string status = a.Status ?? "";
                        bool isOverdue = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                            && a.DueDate.HasValue
                            && a.DueDate.Value.Date < DateTime.Today;

                        bool isWithMe = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                            && !string.IsNullOrWhiteSpace(a.BallHolder)
                            && IsAssignedToUser(a.BallHolder, _currentUserName);

                        return isWithMe && !isOverdue;
                    })
                    .Count();

                OpenActionsCollapsedSummary.Text = $"Overdue: {_overdueActions.Count}, With me: {myActionsCount}";
            }
        }

        // Called from ThisAddIn when email selection changes
        // NOTE: ThisAddIn already wraps this call in Dispatcher.Invoke, so we don't need to do it again
        // PERFORMANCE: Now accepts pre-extracted EmailProperties to avoid blocking UI thread with COM calls
        public async void OnEmailSelected(EmailProperties emailProps)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            System.Diagnostics.Debug.WriteLine($"OnEmailSelected called with: {emailProps?.Subject ?? "(null)"}");

            // NEW: If compose mode is active, exit it first (user switched to regular email)
            if (_isComposeMode)
            {
                System.Diagnostics.Debug.WriteLine("  Exiting compose mode - user selected regular email");
                OnComposeItemDeactivated();
            }

            // Cache both the properties and the MailItem reference
            _currentEmailProperties = emailProps;
            _currentMail = emailProps?.MailItem;
            _isExpanded = false;  // Reset expansion on email change

            if (emailProps == null)
            {
                ActionComboBox.ItemsSource = null;
                ClearLinkedActionFields();
                UpdateButtonStates();
                return;
            }

            // Use pre-extracted identifiers (no PropertyAccessor calls needed!)
            string internetMessageId = emailProps.InternetMessageId;
            string inReplyToId = emailProps.InReplyToId;
            string conversationId = emailProps.ConversationId;

            sw.Stop();
            System.Diagnostics.Debug.WriteLine($"  OnEmailSelected initial setup took {sw.ElapsedMilliseconds}ms on UI thread");

            // Initial grouping (before async classification)
            var initialGrouping = _groupingService.GroupActions(
                _openActions,
                internetMessageId,
                inReplyToId,
                conversationId,
                null,  // No package context yet
                null   // No project context yet
            );

            // Check if linked actions provide context
            if (initialGrouping.LinkedActions.Count > 0)
            {
                _currentPackageContext = initialGrouping.DetectedPackage;
                _currentProjectContext = initialGrouping.DetectedProject;
                RefreshDropdown(initialGrouping);
                AutoSelectBestAction(initialGrouping);
            }
            else
            {
                // Show initial dropdown without package/project context
                _currentPackageContext = null;
                _currentProjectContext = null;
                RefreshDropdown(initialGrouping);

                // Start async classification (uses cached properties)
                await ClassifyEmailContextAsync();
            }

            UpdateButtonStates();
            System.Diagnostics.Debug.WriteLine($"Email selection updated: {emailProps.Subject}");
        }

        private string GetInternetMessageId(Outlook.MailItem mail)
        {
            if (mail == null) return string.Empty;

            const string PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E";

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
            if (mail == null) return string.Empty;

            const string PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001E";

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

        // PERFORMANCE: Now uses cached _currentEmailProperties instead of making redundant COM calls
        private async Task ClassifyEmailContextAsync()
        {
            if (_currentEmailProperties == null)
                return;

            try
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();

                // Use cached properties (already extracted on COM thread)
                string subject = _currentEmailProperties.Subject ?? "";
                string body = _currentEmailProperties.Body ?? "";
                string sender = _currentEmailProperties.SenderEmailAddress ?? "";
                string to = _currentEmailProperties.To ?? "";

                // Run classification on background thread to avoid blocking UI
                var classification = await Task.Run(() =>
                    _classifierService.Classify(subject, body, sender, to));

                sw.Stop();
                System.Diagnostics.Debug.WriteLine($"  Classification took {sw.ElapsedMilliseconds}ms on background thread");

                _currentPackageContext = classification.SuggestedPackageID;
                _currentProjectContext = classification.SuggestedProjectID;

                // Use cached identifiers (no redundant PropertyAccessor calls!)
                string internetMessageId = _currentEmailProperties.InternetMessageId;
                string inReplyToId = _currentEmailProperties.InReplyToId;
                string conversationId = _currentEmailProperties.ConversationId;

                var grouping = _groupingService.GroupActions(
                    _openActions,
                    internetMessageId,
                    inReplyToId,
                    conversationId,
                    _currentPackageContext,
                    _currentProjectContext
                );

                Dispatcher.Invoke(() =>
                {
                    RefreshDropdown(grouping);
                    AutoSelectBestAction(grouping);
                    UpdateButtonStates();
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Classification error: {ex.Message}");
            }
        }

        private void RefreshDropdown(ActionGroupingResult grouping)
        {
            _dropdownItems = BuildDropdownItems(grouping, _isExpanded);
            ActionComboBox.ItemsSource = _dropdownItems;
        }

        private void AutoSelectBestAction(ActionGroupingResult grouping)
        {
            // Only auto-select if there are linked actions
            // Otherwise leave unselected so user can manually choose
            if (grouping.LinkedActions.Count > 0)
            {
                var bestMatch = grouping.LinkedActions[0];
                var item = _dropdownItems.FirstOrDefault(d =>
                    d.ItemType == ActionDropdownItemType.Action &&
                    d.Action?.Id == bestMatch.Id);
                ActionComboBox.SelectedItem = item;
            }
            else
            {
                // No linked actions - clear selection
                ActionComboBox.SelectedItem = null;
            }
        }

        private List<ActionDropdownItem> BuildDropdownItems(ActionGroupingResult grouping, bool expanded)
        {
            var items = new List<ActionDropdownItem>();

            // Section 1: Linked Actions or "No linked actions" message
            if (grouping.LinkedActions.Count > 0)
            {
                items.Add(new ActionDropdownItem
                {
                    ItemType = ActionDropdownItemType.Header,
                    HeaderText = "LINKED ACTIONS"
                });

                foreach (var action in grouping.LinkedActions)
                {
                    items.Add(new ActionDropdownItem
                    {
                        ItemType = ActionDropdownItemType.Action,
                        Action = action,
                        Category = "Linked"
                    });
                }
            }
            else
            {
                // Show "No linked actions" message
                items.Add(new ActionDropdownItem
                {
                    ItemType = ActionDropdownItemType.Header,
                    HeaderText = "No linked actions - select one to update"
                });
            }

            // Section 2: Package Actions
            if (grouping.PackageActions.Count > 0)
            {
                string packageHeader = string.IsNullOrEmpty(grouping.DetectedPackage)
                    ? "PACKAGE ACTIONS"
                    : $"PACKAGE: {grouping.DetectedPackage}";

                items.Add(new ActionDropdownItem
                {
                    ItemType = ActionDropdownItemType.Header,
                    HeaderText = packageHeader
                });

                foreach (var action in grouping.PackageActions)
                {
                    items.Add(new ActionDropdownItem
                    {
                        ItemType = ActionDropdownItemType.Action,
                        Action = action,
                        Category = "Package"
                    });
                }
            }

            // More... expander (if there are project/other actions)
            bool hasMoreActions = grouping.ProjectActions.Count > 0 || grouping.OtherActions.Count > 0;

            if (hasMoreActions)
            {
                if (!expanded)
                {
                    items.Add(new ActionDropdownItem
                    {
                        ItemType = ActionDropdownItemType.MoreExpander
                    });
                }
                else
                {
                    // Expanded: Show Project and Other sections
                    if (grouping.ProjectActions.Count > 0)
                    {
                        string projectHeader = string.IsNullOrEmpty(grouping.DetectedProject)
                            ? "PROJECT ACTIONS"
                            : $"PROJECT: {grouping.DetectedProject}";

                        items.Add(new ActionDropdownItem
                        {
                            ItemType = ActionDropdownItemType.Header,
                            HeaderText = projectHeader
                        });

                        foreach (var action in grouping.ProjectActions)
                        {
                            items.Add(new ActionDropdownItem
                            {
                                ItemType = ActionDropdownItemType.Action,
                                Action = action,
                                Category = "Project"
                            });
                        }
                    }

                    if (grouping.OtherActions.Count > 0)
                    {
                        items.Add(new ActionDropdownItem
                        {
                            ItemType = ActionDropdownItemType.Header,
                            HeaderText = "OTHER"
                        });

                        foreach (var action in grouping.OtherActions)
                        {
                            items.Add(new ActionDropdownItem
                            {
                                ItemType = ActionDropdownItemType.Action,
                                Action = action,
                                Category = "Other"
                            });
                        }
                    }
                }
            }

            return items;
        }

        private void ActionComboBox_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Check if clicking on "More..." expander
            var hitTest = e.OriginalSource as FrameworkElement;
            if (hitTest?.DataContext is ActionDropdownItem item &&
                item.ItemType == ActionDropdownItemType.MoreExpander)
            {
                e.Handled = true;  // Prevent selection

                _isExpanded = true;

                // Re-group and refresh using cached properties
                string internetMessageId = _currentEmailProperties?.InternetMessageId ?? "";
                string inReplyToId = _currentEmailProperties?.InReplyToId ?? "";
                string conversationId = _currentEmailProperties?.ConversationId ?? "";

                var grouping = _groupingService.GroupActions(
                    _openActions,
                    internetMessageId,
                    inReplyToId,
                    conversationId,
                    _currentPackageContext,
                    _currentProjectContext
                );

                RefreshDropdown(grouping);

                // Keep dropdown open
                ActionComboBox.IsDropDownOpen = true;
            }
        }

        private void ActionComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;

            // Ignore non-action selections
            if (selectedItem == null || selectedItem.ItemType != ActionDropdownItemType.Action)
            {
                if (selectedItem?.ItemType == ActionDropdownItemType.Header ||
                    selectedItem?.ItemType == ActionDropdownItemType.MoreExpander)
                {
                    // Revert to previous selection
                    e.Handled = true;
                    return;
                }

                // No action selected - clear fields (visibility controlled by XAML)
                ClearLinkedActionFields();
                RelatedMessagesListView.ItemsSource = null;
                UpdateButtonStates();
                return;
            }

            // Action selected - populate fields (visibility controlled by XAML)
            var selectedAction = selectedItem.Action;
            PopulateLinkedActionFields(selectedAction);
            LoadRelatedMessagesAsync(selectedAction);

            // Store previous values for change detection
            _previousActionState = new ActionFieldState
            {
                Project = selectedAction.Project,
                Package = selectedAction.Package,
                Title = selectedAction.Title,
                BallHolder = selectedAction.BallHolder,
                DueDate = selectedAction.DueDate
            };

            UpdateButtonStates();
        }

        /// <summary>
        /// Updates the health indicator dots based on action status, due date, and assignment.
        /// </summary>
        /// <param name="action">The action to evaluate</param>
        private void UpdateHealthIndicators(ActionItem action)
        {
            if (action == null)
            {
                // Hide both indicators when no action
                OverdueIndicator.Visibility = Visibility.Collapsed;
                WithMeIndicator.Visibility = Visibility.Collapsed;
                return;
            }

            // Check if action is NOT Closed (indicators apply to all non-closed statuses)
            bool isNotClosed = !string.Equals(action.Status, "Closed", StringComparison.OrdinalIgnoreCase);

            // OVERDUE CHECK
            // Show red dot if: Status != "Closed" AND DueDate < Today (strictly past)
            bool isOverdue = false;
            if (isNotClosed && action.DueDate.HasValue)
            {
                isOverdue = action.DueDate.Value.Date < DateTime.Today;
            }
            OverdueIndicator.Visibility = isOverdue ? Visibility.Visible : Visibility.Collapsed;

            // ASSIGNMENT CHECK ("With Me")
            // Show amber dot if: Status != "Closed" AND BallHolder contains user's name (flexible matching)
            bool isWithMe = false;
            if (isNotClosed && !string.IsNullOrWhiteSpace(action.BallHolder) && !string.IsNullOrWhiteSpace(_currentUserName))
            {
                isWithMe = IsAssignedToUser(action.BallHolder, _currentUserName);
            }
            WithMeIndicator.Visibility = isWithMe ? Visibility.Visible : Visibility.Collapsed;

            // Debug output
            System.Diagnostics.Debug.WriteLine($"Health Indicators - Action: {action.Title}, Status: {action.Status}, " +
                $"Overdue: {isOverdue}, WithMe: {isWithMe}");
        }

        /// <summary>
        /// Checks if the BallHolder field indicates assignment to the current user.
        /// Supports flexible name matching: Full name, First name, Last name, or partial match.
        /// </summary>
        /// <param name="ballHolder">The BallHolder field value (e.g., "Wally C", "Cloud, Wally", "Wally Cloud")</param>
        /// <param name="userName">The user's configured name (e.g., "Wally Cloud")</param>
        /// <returns>True if the BallHolder is assigned to the user</returns>
        private bool IsAssignedToUser(string ballHolder, string userName)
        {
            // Normalize for comparison (case-insensitive, trim whitespace)
            string ballHolderNormalized = ballHolder.Trim().ToLowerInvariant();
            string userNameNormalized = userName.Trim().ToLowerInvariant();

            // Check 1: Exact match (handles "Wally Cloud" == "Wally Cloud")
            if (ballHolderNormalized.Contains(userNameNormalized))
                return true;

            // Split user name into parts for flexible matching
            string[] userParts = userNameNormalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (userParts.Length == 0)
                return false;

            // Check 2: Full name match in reverse ("Cloud, Wally" or "Cloud Wally")
            if (userParts.Length >= 2)
            {
                string reversedName = $"{userParts[userParts.Length - 1]} {userParts[0]}"; // "cloud wally"
                if (ballHolderNormalized.Contains(reversedName))
                    return true;

                string reversedNameWithComma = $"{userParts[userParts.Length - 1]}, {userParts[0]}"; // "cloud, wally"
                if (ballHolderNormalized.Contains(reversedNameWithComma))
                    return true;
            }

            // Check 3: First name or Last name match (handles "Wally", "Cloud", "Wally C")
            foreach (string part in userParts)
            {
                if (part.Length >= 2 && ballHolderNormalized.Contains(part))
                    return true;
            }

            // Check 4: Initial matching (handles "W. Cloud", "Wally C.", "W.C.")
            // For each part, check if BallHolder contains first letter + period
            foreach (string part in userParts)
            {
                if (part.Length > 0)
                {
                    string initial = $"{part[0]}.";
                    if (ballHolderNormalized.Contains(initial))
                        return true;
                }
            }

            return false;
        }

        private void PopulateLinkedActionFields(ActionItem action)
        {
            if (action == null)
            {
                ClearLinkedActionFields();
                return;
            }

            Dispatcher.Invoke(() =>
            {
                LinkedProjectTextBox.Text = action.Project ?? "";
                LinkedPackageTextBox.Text = action.Package ?? "";
                LinkedTitleTextBox.Text = action.Title ?? "";
                LinkedBallHolderTextBox.Text = action.BallHolder ?? "";
                LinkedDueDatePicker.SelectedDate = action.DueDate;

                // Update health indicators for the newly displayed action
                UpdateHealthIndicators(action);
            });
        }

        private void ClearLinkedActionFields()
        {
            Dispatcher.Invoke(() =>
            {
                LinkedProjectTextBox.Text = "";
                LinkedPackageTextBox.Text = "";
                LinkedTitleTextBox.Text = "";
                LinkedBallHolderTextBox.Text = "";
                LinkedDueDatePicker.SelectedDate = null;
                RelatedMessagesListView.ItemsSource = null;

                // Hide health indicators when no action is displayed
                UpdateHealthIndicators(null);
            });

            _previousActionState = null;
        }

        private async void SaveLinkedActionButton_Click(object sender, RoutedEventArgs e)
        {
            await SaveLinkedActionFieldsAsync();
        }

        private async void RefreshDataButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SetStatus("Refreshing data from Google Sheets...");
                RefreshDataButton.IsEnabled = false;

                // Reload classification rules and actions
                await Task.Run(async () =>
                {
                    var rulesData = await _googleSheetsService.FetchConfigRulesAsync();
                    _classifierService.LoadRules(rulesData);
                    _openActions = await _googleSheetsService.FetchOpenActionsAsync();
                });

                // Update UI
                Dispatcher.Invoke(() =>
                {
                    // Re-process current email if one is selected (will trigger grouping)
                    if (_currentEmailProperties != null)
                    {
                        OnEmailSelected(_currentEmailProperties);
                    }
                    else
                    {
                        ActionComboBox.ItemsSource = null;
                    }

                    RefreshDashboard();
                });

                SetStatus($"Data refreshed successfully. Loaded {_classifierService.GetRuleCount()} rules and {_openActions.Count} open actions.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Refresh error: {ex.Message}");
                SetStatus($"Error refreshing data: {ex.Message}");
                MessageBox.Show($"Error refreshing data from Google Sheets:\n{ex.Message}", "Refresh Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                RefreshDataButton.IsEnabled = true;
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var settingsWindow = new SettingsWindow();
                bool? result = settingsWindow.ShowDialog();

                if (result == true)
                {
                    // Settings were saved, reload current user name
                    _currentUserName = Properties.Settings.Default.UserName;

                    // Refresh dashboard in case WithMe indicators need updating
                    RefreshDashboard();

                    SetStatus("Settings saved successfully.");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Settings error: {ex.Message}");
                MessageBox.Show($"Error opening settings:\n{ex.Message}", "Settings Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void LinkedField_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                e.Handled = true;
                await SaveLinkedActionFieldsAsync();
            }
        }

        private async Task SaveLinkedActionFieldsAsync()
        {
            var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
            var selectedAction = selectedItem?.Action;
            if (selectedAction == null)
            {
                MessageBox.Show("No action selected.", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                SetStatus("Saving linked action changes...");

                // Disable controls during save
                SaveLinkedActionButton.IsEnabled = false;
                LinkedActionBorder.IsEnabled = false;

                // Get values from UI
                string project = LinkedProjectTextBox.Text?.Trim() ?? "";
                string package = LinkedPackageTextBox.Text?.Trim() ?? "";
                string title = LinkedTitleTextBox.Text?.Trim() ?? "";
                string ballHolder = LinkedBallHolderTextBox.Text?.Trim() ?? "";
                DateTime? dueDate = LinkedDueDatePicker.SelectedDate;

                // Validate
                if (string.IsNullOrWhiteSpace(title))
                {
                    MessageBox.Show("Title cannot be empty.", "Validation Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Call Google Sheets service
                await _googleSheetsService.UpdateActionFieldsAsync(
                    selectedAction.Id,
                    project,
                    package,
                    title,
                    ballHolder,
                    dueDate
                );

                // Update local object
                selectedAction.Project = project;
                selectedAction.Package = package;
                selectedAction.Title = title;
                selectedAction.BallHolder = ballHolder;
                selectedAction.DueDate = dueDate;

                // Refresh health indicators after save (status/assignment/due date may have changed)
                Dispatcher.Invoke(() => UpdateHealthIndicators(selectedAction));

                // Refresh ComboBox display - rebuild with grouping using cached properties
                if (_currentEmailProperties != null)
                {
                    string internetMessageId = _currentEmailProperties.InternetMessageId;
                    string inReplyToId = _currentEmailProperties.InReplyToId;
                    string conversationId = _currentEmailProperties.ConversationId;
                    var grouping = _groupingService.GroupActions(
                        _openActions,
                        internetMessageId,
                        inReplyToId,
                        conversationId,
                        _currentPackageContext,
                        _currentProjectContext
                    );
                    RefreshDropdown(grouping);

                    // Re-select the updated action
                    var item = _dropdownItems.FirstOrDefault(d =>
                        d.ItemType == ActionDropdownItemType.Action &&
                        d.Action?.Id == selectedAction.Id);
                    ActionComboBox.SelectedItem = item;
                }

                SetStatus("Linked action saved successfully");
            }
            catch (Exception ex)
            {
                SetStatus($"Error saving: {ex.Message}");
                MessageBox.Show($"Error saving linked action: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // Re-enable controls
                SaveLinkedActionButton.IsEnabled = true;
                LinkedActionBorder.IsEnabled = true;
            }
        }

        private async Task DetectAndHighlightChangesAsync(ActionItem updatedAction)
        {
            if (_previousActionState == null)
                return;

            var changedFields = new List<TextBox>();

            // Detect which fields changed
            if (updatedAction.Project != _previousActionState.Project)
                changedFields.Add(LinkedProjectTextBox);

            if (updatedAction.Package != _previousActionState.Package)
                changedFields.Add(LinkedPackageTextBox);

            if (updatedAction.Title != _previousActionState.Title)
                changedFields.Add(LinkedTitleTextBox);

            if (updatedAction.BallHolder != _previousActionState.BallHolder)
                changedFields.Add(LinkedBallHolderTextBox);

            if (changedFields.Count == 0)
                return;

            // Animate changed fields
            await Dispatcher.InvokeAsync(async () =>
            {
                var highlightColor = new SolidColorBrush(Color.FromRgb(204, 255, 204)); // Pastel green
                var originalColor = new SolidColorBrush(Colors.White);

                // Set highlight
                foreach (var field in changedFields)
                {
                    field.Background = highlightColor;
                }

                // Wait 2 seconds
                await Task.Delay(2000);

                // Fade back
                foreach (var field in changedFields)
                {
                    field.Background = originalColor;
                }
            });
        }

        private void UpdateButtonStates()
        {
            var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
            bool hasSelection = selectedItem?.ItemType == ActionDropdownItemType.Action;
            bool hasEmail = _currentMail != null || _composeMail != null;

            if (_isComposeMode)
            {
                // In compose mode: Use UpdateButtonVisualsForSchedule which handles all button states
                CreateButton.IsEnabled = hasEmail;
                // Other buttons handled by UpdateButtonVisualsForSchedule
                UpdateButtonVisualsForSchedule();
            }
            else
            {
                // Normal mode: Original logic
                CreateButton.IsEnabled = hasEmail;
                UpdateButton.IsEnabled = hasSelection && hasEmail;
                CreateMultipleButton.IsEnabled = hasEmail;
                CloseButton.IsEnabled = hasSelection;
                ReopenButton.IsEnabled = hasSelection;
            }
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"CreateButton_Click: _isComposeMode={_isComposeMode}");

                if (_isComposeMode)
                {
                    // Toggle scheduled create
                    if (_currentDeferredData?.Mode == "Create")
                    {
                        System.Diagnostics.Debug.WriteLine("  Canceling Create schedule");
                        // Cancel scheduling
                        ClearDeferredData(_composeMail);
                        _currentDeferredData = null;
                        SetStatus("Create action schedule canceled");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("  Scheduling Create action");
                        // Schedule create - MUTUAL EXCLUSIVITY: clear Update schedule first
                        if (_currentDeferredData?.Mode == "Update")
                        {
                            System.Diagnostics.Debug.WriteLine("  Clearing existing Update schedule first");
                            ClearDeferredData(_composeMail);
                        }

                        var data = new DeferredActionData { Mode = "Create" };
                        SaveDeferredData(_composeMail, data);
                        _currentDeferredData = data;
                        SetStatus("Create action scheduled - will execute after send");
                    }
                    UpdateButtonVisualsForSchedule();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  Normal mode - immediate create");
                    // Normal mode - immediate create
                    if (_currentMail == null)
                    {
                        MessageBox.Show("Please select an email first.", "No Email",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // Capture the current mail before queuing
                    var mailToProcess = _currentMail;

                    // Queue the action
                    _actionQueue.Enqueue(() => ProcessCreateActionAsync(mailToProcess));
                    ProcessQueue();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in CreateButton_Click: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"Error: {ex.Message}", "Create Button Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreateMultipleButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"CreateMultipleButton_Click: _isComposeMode={_isComposeMode}");

                if (_isComposeMode)
                {
                    // Toggle scheduled create multiple
                    if (_currentDeferredData?.Mode == "CreateMultiple")
                    {
                        System.Diagnostics.Debug.WriteLine("  Canceling CreateMultiple schedule");
                        // Cancel scheduling
                        ClearDeferredData(_composeMail);
                        _currentDeferredData = null;
                        SetStatus("Create multiple actions schedule canceled");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("  Scheduling CreateMultiple action");
                        // Clear any existing schedule first
                        if (_currentDeferredData != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"  Clearing existing {_currentDeferredData.Mode} schedule first");
                            ClearDeferredData(_composeMail);
                        }

                        var data = new DeferredActionData { Mode = "CreateMultiple" };
                        SaveDeferredData(_composeMail, data);
                        _currentDeferredData = data;
                        SetStatus("Create multiple actions scheduled - will execute after send");
                    }
                    UpdateButtonVisualsForSchedule();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  Normal mode - immediate create multiple");
                    // Normal mode - immediate create multiple
                    if (_currentMail == null)
                    {
                        MessageBox.Show("Please select an email first.", "No Email",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // Capture the current mail before queuing
                    var mailToProcess = _currentMail;

                    // Queue the action
                    _actionQueue.Enqueue(() => ProcessCreateMultipleActionsAsync(mailToProcess));
                    ProcessQueue();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in CreateMultipleButton_Click: {ex.Message}");
                MessageBox.Show($"Error: {ex.Message}", "Create Multiple Button Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ProcessCreateActionAsync(Outlook.MailItem mail)
        {
            try
            {
                SetStatus("Processing action...");

                // Get email data
                string body = mail.Body ?? "";
                string senderEmail = mail.SenderEmailAddress ?? "";
                string toRecipients = mail.To ?? "";
                string subject = mail.Subject ?? "";

                // AUTO-CLASSIFY Project and Package
                // Check both FROM sender and TO recipients
                var classification = _classifierService.Classify(subject, body, senderEmail, toRecipients);

                // Handle ambiguity
                if (classification.IsAmbiguous)
                {
                    // Show ambiguity resolution dialog
                    var ambiguityDialog = new AmbiguityResolutionDialog
                    {
                        AmbiguityReason = classification.AmbiguityReason,
                        Candidates = classification.Candidates
                    };

                    if (ambiguityDialog.ShowDialog() == true)
                    {
                        // User selected from candidates
                        var selectedProject = ambiguityDialog.SelectedCandidates
                            .FirstOrDefault(c => c.Type == "PROJECT");
                        var selectedPackage = ambiguityDialog.SelectedCandidates
                            .FirstOrDefault(c => c.Type == "PACKAGE");

                        classification.SuggestedProjectID = selectedProject?.Name ?? classification.SuggestedProjectID;
                        classification.SuggestedPackageID = selectedPackage?.Name ?? classification.SuggestedPackageID;
                    }
                    else
                    {
                        // User cancelled - use defaults
                        SetStatus("Classification cancelled");
                        return;
                    }
                }

                string project = classification.SuggestedProjectID;
                string package = classification.SuggestedPackageID;

                // Get LLM suggestions for Title, BallHolder, Description
                LLMExtractionResult extraction;
                try
                {
                    extraction = await _llmService.GetExtractionAsync(body, senderEmail, subject);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LLM extraction failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsLLMError(mail);
                    throw;
                }

                bool shouldCreate = true;
                string title = extraction.Title;
                string ballHolder = extraction.BallHolder;
                string description = extraction.Description;

                // Check if confirmation is enabled
                if (Properties.Settings.Default.ConfirmActions)
                {
                    SetStatus("Waiting for confirmation...");

                    // Show dialog for confirmation (WITH PROJECT/PACKAGE)
                    var dialog = new CreateActionDialog
                    {
                        Project = project,
                        Package = package,
                        ActionTitle = title,
                        BallHolder = ballHolder,
                        Description = description
                    };

                    shouldCreate = dialog.ShowDialog() == true;

                    if (shouldCreate)
                    {
                        // User may have edited the values
                        project = dialog.Project;
                        package = dialog.Package;
                        title = dialog.ActionTitle;
                        ballHolder = dialog.BallHolder;
                        description = dialog.Description;
                    }
                }

                if (shouldCreate)
                {
                    SetStatus("Creating action in Google Sheets...");

                    string conversationId = mail.ConversationID;
                    string emailReference = GetEmailReference(mail);
                    DateTime emailDate = GetEmailDate(mail);
                    int dueDays = GetDefaultDueDays();

                    try
                    {
                        await _googleSheetsService.CreateActionAsync(
                            project,        // NEW parameter
                            package,        // NEW parameter
                            title,
                            ballHolder,
                            conversationId,
                            emailReference,
                            description,
                            emailDate,
                            dueDays
                        );

                        // Tag email as tracked (fire-and-forget)
                        _emailCategoryService.MarkEmailAsTracked(mail);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Google Sheets CreateAction failed: {ex.Message}");
                        _emailCategoryService.MarkEmailAsSheetsError(mail);
                        throw;
                    }

                    SetStatus($"Created: {title}");
                    LoadActionsAsync();
                }
                else
                {
                    SetStatus("Action cancelled");
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Error: {ex.Message}");
                MessageBox.Show($"Error creating action: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ProcessCreateMultipleActionsAsync(Outlook.MailItem mail)
        {
            try
            {
                SetStatus("Extracting multiple actions from email...");

                // Extract email data
                string body = mail.Body ?? "";
                string senderEmail = mail.SenderEmailAddress ?? "";
                string toRecipients = mail.To ?? "";
                string subject = mail.Subject ?? "";

                // Auto-classify Project & Package (shared by all actions)
                SetStatus("Auto-classifying Project and Package...");
                var classification = _classifierService.Classify(subject, body, senderEmail, toRecipients);

                string suggestedProject = classification.SuggestedProjectID ?? "Random";
                string suggestedPackage = classification.SuggestedPackageID ?? "";

                // Handle classification ambiguity (show disambiguation dialog)
                if (classification.IsAmbiguous)
                {
                    var ambiguityDialog = new AmbiguityResolutionDialog
                    {
                        AmbiguityReason = classification.AmbiguityReason,
                        Candidates = classification.Candidates
                    };

                    if (ambiguityDialog.ShowDialog() == true)
                    {
                        // User selected from candidates
                        var selectedProject = ambiguityDialog.SelectedCandidates
                            .FirstOrDefault(c => c.Type == "PROJECT");
                        var selectedPackage = ambiguityDialog.SelectedCandidates
                            .FirstOrDefault(c => c.Type == "PACKAGE");

                        suggestedProject = selectedProject?.Name ?? suggestedProject;
                        suggestedPackage = selectedPackage?.Name ?? suggestedPackage;
                    }
                    else
                    {
                        SetStatus("Action creation cancelled");
                        return;
                    }
                }

                // Call LLM to extract MULTIPLE actions
                SetStatus("Analyzing email for multiple actions...");
                List<LLMExtractionResult> extractions;
                try
                {
                    extractions = await _llmService.GetMultipleExtractionsAsync(body, senderEmail, subject);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LLM multiple extractions failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsLLMError(mail);
                    throw;
                }

                if (extractions == null || extractions.Count == 0)
                {
                    MessageBox.Show("No actions found in this email.\n\nThe email may not contain clear action items.",
                        "No Actions Found",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    SetStatus("No actions found");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"LLM extracted {extractions.Count} actions");

                // Show confirmation dialog if enabled
                bool confirmActions = Properties.Settings.Default.ConfirmActions;

                int successCount = 0;
                int failedCount = 0;
                List<string> createdTitles = new List<string>();

                // Get shared properties
                string conversationId = mail.ConversationID;
                string emailReference = GetEmailReference(mail);
                DateTime sentOn = GetEmailDate(mail);
                int defaultDueDays = GetDefaultDueDays();

                // Create each action
                for (int i = 0; i < extractions.Count; i++)
                {
                    var extraction = extractions[i];
                    int actionNumber = i + 1;

                    try
                    {
                        SetStatus($"Creating action {actionNumber} of {extractions.Count}: {extraction.Title}");

                        string finalTitle = extraction.Title;
                        string finalBallHolder = extraction.BallHolder;
                        string finalDescription = extraction.Description;
                        string finalProject = suggestedProject;
                        string finalPackage = suggestedPackage;

                        // Show confirmation dialog for THIS action if enabled
                        if (confirmActions)
                        {
                            var confirmDialog = new CreateActionDialog
                            {
                                Project = finalProject,
                                Package = finalPackage,
                                ActionTitle = finalTitle,
                                BallHolder = finalBallHolder,
                                Description = finalDescription
                            };

                            bool? dialogResult = confirmDialog.ShowDialog();
                            if (dialogResult != true)
                            {
                                System.Diagnostics.Debug.WriteLine($"Action {actionNumber} cancelled by user");
                                failedCount++;
                                continue; // Skip this action
                            }

                            // Get potentially edited values
                            finalProject = confirmDialog.Project;
                            finalPackage = confirmDialog.Package;
                            finalTitle = confirmDialog.ActionTitle;
                            finalBallHolder = confirmDialog.BallHolder;
                            finalDescription = confirmDialog.Description;
                        }

                        // Create action in Google Sheets
                        string initialNote = finalDescription;

                        int newActionId = await _googleSheetsService.CreateActionAsync(
                            finalProject,
                            finalPackage,
                            finalTitle,
                            finalBallHolder,
                            conversationId,
                            emailReference,
                            initialNote,
                            sentOn,
                            defaultDueDays
                        );

                        // Tag email as tracked after each action creation (fire-and-forget)
                        _emailCategoryService.MarkEmailAsTracked(mail);

                        System.Diagnostics.Debug.WriteLine($"Created action {actionNumber}: ID={newActionId}, Title={finalTitle}");
                        successCount++;
                        createdTitles.Add(finalTitle);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error creating action {actionNumber}: {ex.Message}");
                        // Could be LLM error or Sheets error - mark both to ensure visibility
                        _emailCategoryService.MarkEmailAsLLMError(mail);
                        _emailCategoryService.MarkEmailAsSheetsError(mail);
                        failedCount++;
                    }
                }

                // Reload actions from Google Sheets
                await LoadActionsAsync();

                // Show summary
                string summary = $"Successfully created {successCount} action(s)";
                if (failedCount > 0)
                    summary += $"\n{failedCount} action(s) failed or were cancelled";

                if (createdTitles.Count > 0)
                {
                    summary += "\n\nCreated actions:\n";
                    for (int i = 0; i < createdTitles.Count; i++)
                    {
                        summary += $"{i + 1}. {createdTitles[i]}\n";
                    }
                }

                MessageBox.Show(summary, "Multiple Actions Created",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                SetStatus($"Created {successCount} action(s) successfully");
            }
            catch (Exception ex)
            {
                SetStatus("Error creating multiple actions");
                MessageBox.Show($"Error creating actions:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"UpdateButton_Click: _isComposeMode={_isComposeMode}");

                if (_isComposeMode)
                {
                    // Toggle scheduled update
                    if (_currentDeferredData?.Mode == "Update")
                    {
                        System.Diagnostics.Debug.WriteLine("  Canceling Update schedule");
                        // Cancel scheduling
                        ClearDeferredData(_composeMail);
                        _currentDeferredData = null;
                        SetStatus("Update action schedule canceled");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("  Scheduling Update action");
                        // Schedule update - must have action selected
                        var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
                        var selectedAction = selectedItem?.Action;
                        if (selectedAction == null)
                        {
                            System.Diagnostics.Debug.WriteLine("  No action selected - showing warning");
                            MessageBox.Show("Please select an action to update.", "No Action Selected",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        // MUTUAL EXCLUSIVITY: clear Create schedule first
                        if (_currentDeferredData?.Mode == "Create")
                        {
                            System.Diagnostics.Debug.WriteLine("  Clearing existing Create schedule first");
                            ClearDeferredData(_composeMail);
                        }

                        var data = new DeferredActionData
                        {
                            Mode = "Update",
                            ActionID = selectedAction.Id
                        };
                        SaveDeferredData(_composeMail, data);
                        _currentDeferredData = data;
                        SetStatus($"Update action '{selectedAction.Title}' scheduled - will execute after send");
                    }
                    UpdateButtonVisualsForSchedule();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  Normal mode - immediate update");
                    // Normal mode - immediate update
                    var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
                    var selectedAction = selectedItem?.Action;
                    if (selectedAction == null || _currentMail == null)
                        return;

                    // Capture current values before queuing
                    var mailToProcess = _currentMail;
                    var actionToUpdate = selectedAction;

                    // Queue the action
                    _actionQueue.Enqueue(() => ProcessUpdateActionAsync(mailToProcess, actionToUpdate));
                    ProcessQueue();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in UpdateButton_Click: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"Error: {ex.Message}", "Update Button Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateButtonVisualsForSchedule()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"UpdateButtonVisualsForSchedule: _isComposeMode={_isComposeMode}, DeferredMode={_currentDeferredData?.Mode ?? "None"}");

                // Show/hide compose mode header
                ComposeModeHeader.Visibility = _isComposeMode ? Visibility.Visible : Visibility.Collapsed;

                if (_isComposeMode)
                {
                    // Compose mode - enable buttons based on context
                    System.Diagnostics.Debug.WriteLine("  Compose mode - updating button states");

                    // Check if an action is selected
                    var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
                    var hasActionSelected = selectedItem?.Action != null;
                    System.Diagnostics.Debug.WriteLine($"  Action selected: {hasActionSelected}");

                    // Create and Create Multiple are always available in compose mode
                    CreateMultipleButton.IsEnabled = true;

                    // Update, Close, Reopen require an action to be selected
                    UpdateButton.IsEnabled = hasActionSelected;
                    CloseButton.IsEnabled = hasActionSelected;
                    ReopenButton.IsEnabled = hasActionSelected;

                    // Update button visuals based on scheduled action
                    var primaryBrush = (SolidColorBrush)FindResource("PrimaryAccent");
                    var secondaryBrush = (SolidColorBrush)FindResource("SecondaryButton");

                    // Create button
                    if (_currentDeferredData?.Mode == "Create")
                    {
                        CreateButton.Background = ScheduledBrush;
                        CreateButton.Content = "✓ Cancel Create";
                    }
                    else
                    {
                        CreateButton.Background = primaryBrush;
                        CreateButton.Content = "Create New";
                    }

                    // Create Multiple button (now primary blue)
                    if (_currentDeferredData?.Mode == "CreateMultiple")
                    {
                        CreateMultipleButton.Background = ScheduledBrush;
                        CreateMultipleButton.Content = "✓ Cancel Multiple";
                    }
                    else
                    {
                        CreateMultipleButton.Background = primaryBrush;
                        CreateMultipleButton.Content = "Create Multiple";
                    }

                    // Update button
                    if (_currentDeferredData?.Mode == "Update")
                    {
                        UpdateButton.Background = ScheduledBrush;
                        UpdateButton.Content = "✓ Cancel Update";
                    }
                    else
                    {
                        UpdateButton.Background = primaryBrush;
                        UpdateButton.Content = "Update";
                    }

                    // Close button (now primary blue)
                    if (_currentDeferredData?.Mode == "Close")
                    {
                        CloseButton.Background = ScheduledBrush;
                        CloseButton.Content = "✓ Cancel Close";
                    }
                    else
                    {
                        CloseButton.Background = primaryBrush;
                        CloseButton.Content = "Close";
                    }

                    // Reopen button (now primary blue)
                    if (_currentDeferredData?.Mode == "Reopen")
                    {
                        ReopenButton.Background = ScheduledBrush;
                        ReopenButton.Content = "✓ Cancel Reopen";
                    }
                    else
                    {
                        ReopenButton.Background = primaryBrush;
                        ReopenButton.Content = "Reopen";
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  Normal mode - resetting to defaults");
                    // Normal mode - reset to defaults
                    var primaryBrush = (SolidColorBrush)FindResource("PrimaryAccent");
                    var secondaryBrush = (SolidColorBrush)FindResource("SecondaryButton");

                    CreateButton.Background = primaryBrush;
                    CreateButton.Content = "Create New";

                    CreateMultipleButton.Background = primaryBrush;
                    CreateMultipleButton.Content = "Create Multiple";

                    UpdateButton.Background = primaryBrush;
                    UpdateButton.Content = "Update";

                    CloseButton.Background = primaryBrush;
                    CloseButton.Content = "Close";

                    ReopenButton.Background = primaryBrush;
                    ReopenButton.Content = "Reopen";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in UpdateButtonVisualsForSchedule: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        private async Task ProcessUpdateActionAsync(Outlook.MailItem mail, ActionItem selectedAction)
        {
            try
            {
                SetStatus("Processing update...");

                string body = mail.Body ?? "";
                string currentContext = selectedAction.HistoryLog ?? "";
                string currentBallHolder = selectedAction.BallHolder ?? "";

                var delta = await _llmService.GetDeltaAsync(body, currentContext, currentBallHolder);

                bool shouldUpdate = true;
                string ballHolder = delta.NewBallHolder;
                string updateNote = delta.UpdateSummary;

                // Check if confirmation is enabled
                if (Properties.Settings.Default.ConfirmActions)
                {
                    SetStatus("Waiting for confirmation...");

                    var dialog = new UpdateActionDialog
                    {
                        BallHolder = ballHolder,
                        UpdateNote = updateNote
                    };

                    shouldUpdate = dialog.ShowDialog() == true;

                    if (shouldUpdate)
                    {
                        // User may have edited the values
                        ballHolder = dialog.BallHolder;
                        updateNote = dialog.UpdateNote;
                    }
                }

                if (shouldUpdate)
                {
                    SetStatus("Updating action in Google Sheets...");

                    string emailReference = GetEmailReference(mail);
                    DateTime emailDate = GetEmailDate(mail);
                    int dueDays = GetDefaultDueDays();

                    try
                    {
                        await _googleSheetsService.UpdateActionAsync(
                            selectedAction.Id,
                            emailReference,
                            ballHolder,
                            updateNote,
                            emailDate,      // SentOn
                            dueDays         // Due days offset
                        );

                        // Tag email as tracked (fire-and-forget)
                        _emailCategoryService.MarkEmailAsTracked(mail);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Google Sheets UpdateAction failed: {ex.Message}");
                        _emailCategoryService.MarkEmailAsSheetsError(mail);
                        throw;
                    }

                    SetStatus($"Updated: {selectedAction.Title}");

                    // Detect and highlight changed fields
                    await DetectAndHighlightChangesAsync(selectedAction);

                    // Refresh the linked action fields with new data
                    LoadActionsAsync();
                }
                else
                {
                    SetStatus("Update cancelled");
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Error: {ex.Message}");
                MessageBox.Show($"Error updating action: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"CloseButton_Click: _isComposeMode={_isComposeMode}");

                if (_isComposeMode)
                {
                    // Toggle scheduled close
                    if (_currentDeferredData?.Mode == "Close")
                    {
                        System.Diagnostics.Debug.WriteLine("  Canceling Close schedule");
                        // Cancel scheduling
                        ClearDeferredData(_composeMail);
                        _currentDeferredData = null;
                        SetStatus("Close action schedule canceled");
                    }
                    else
                    {
                        // Need an action selected to close
                        var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
                        var selectedAction = selectedItem?.Action;
                        if (selectedAction == null)
                        {
                            MessageBox.Show("Please select an action to close.", "No Action",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        System.Diagnostics.Debug.WriteLine("  Scheduling Close action");
                        // Clear any existing schedule first
                        if (_currentDeferredData != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"  Clearing existing {_currentDeferredData.Mode} schedule first");
                            ClearDeferredData(_composeMail);
                        }

                        var data = new DeferredActionData { Mode = "Close", ActionID = selectedAction.Id };
                        SaveDeferredData(_composeMail, data);
                        _currentDeferredData = data;
                        SetStatus($"Close action '{selectedAction.Title}' scheduled - will execute after send");
                    }
                    UpdateButtonVisualsForSchedule();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  Normal mode - immediate close");
                    // Normal mode - immediate close
                    var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
                    var selectedAction = selectedItem?.Action;
                    if (selectedAction == null || _currentMail == null)
                        return;

                    SetStatus("Generating closure summary...");

                    // Get email details
                    DateTime emailDate = GetEmailDate(_currentMail);
                    string closingEmailReference = GetEmailReference(_currentMail);

                    if (string.IsNullOrWhiteSpace(closingEmailReference))
                    {
                        MessageBox.Show("Could not extract email reference. Please try again.", "Error",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // Generate LLM closure summary
                    string actionContext = $"Title: {selectedAction.Title}\nProject: {selectedAction.Project}\nPackage: {selectedAction.Package}";
                    string closureNote;

                    try
                    {
                        SetStatus("Analyzing email for closure summary...");
                        closureNote = await _llmService.GetClosureSummaryAsync(_currentMail.Body, actionContext);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"LLM closure summary failed: {ex.Message}");
                        _emailCategoryService.MarkEmailAsLLMError(_currentMail);
                        closureNote = "Action closed";
                    }

                    SetStatus("Closing action...");

                    // Close the action with LLM-generated note, email reference, and date
                    try
                    {
                        await _googleSheetsService.CloseActionAsync(
                            selectedAction.Id,
                            closureNote,
                            closingEmailReference,
                            emailDate
                        );

                        // Tag email as tracked (fire-and-forget)
                        _emailCategoryService.MarkEmailAsTracked(_currentMail);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Google Sheets CloseAction failed: {ex.Message}");
                        _emailCategoryService.MarkEmailAsSheetsError(_currentMail);
                        throw;
                    }

                    SetStatus("Action closed successfully");

                    await LoadActionsAsync();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in CloseButton_Click: {ex.Message}");
                SetStatus("Error closing action");
                MessageBox.Show($"Error closing action: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void ReopenButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"ReopenButton_Click: _isComposeMode={_isComposeMode}");

                if (_isComposeMode)
                {
                    // Toggle scheduled reopen
                    if (_currentDeferredData?.Mode == "Reopen")
                    {
                        System.Diagnostics.Debug.WriteLine("  Canceling Reopen schedule");
                        // Cancel scheduling
                        ClearDeferredData(_composeMail);
                        _currentDeferredData = null;
                        SetStatus("Reopen action schedule canceled");
                    }
                    else
                    {
                        // Need an action selected to reopen
                        var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
                        var selectedAction = selectedItem?.Action;
                        if (selectedAction == null)
                        {
                            MessageBox.Show("Please select an action to reopen.", "No Action",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        System.Diagnostics.Debug.WriteLine("  Scheduling Reopen action");
                        // Clear any existing schedule first
                        if (_currentDeferredData != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"  Clearing existing {_currentDeferredData.Mode} schedule first");
                            ClearDeferredData(_composeMail);
                        }

                        var data = new DeferredActionData { Mode = "Reopen", ActionID = selectedAction.Id };
                        SaveDeferredData(_composeMail, data);
                        _currentDeferredData = data;
                        SetStatus($"Reopen action '{selectedAction.Title}' scheduled - will execute after send");
                    }
                    UpdateButtonVisualsForSchedule();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  Normal mode - immediate reopen");
                    // Normal mode - immediate reopen
                    var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
                    var selectedAction = selectedItem?.Action;
                    if (selectedAction == null || _currentMail == null)
                        return;

                    DateTime emailDate = GetEmailDate(_currentMail);
                    int dueDays = GetDefaultDueDays();

                    try
                    {
                        await _googleSheetsService.SetStatusAsync(selectedAction.Id, "Open", emailDate, dueDays);

                        // Tag email as tracked (fire-and-forget)
                        _emailCategoryService.MarkEmailAsTracked(_currentMail);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Google Sheets SetStatus failed: {ex.Message}");
                        _emailCategoryService.MarkEmailAsSheetsError(_currentMail);
                        throw;
                    }

                    MessageBox.Show("Action reopened.", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    LoadActionsAsync();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in ReopenButton_Click: {ex.Message}");
                MessageBox.Show($"Error reopening action: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Extract complete email reference in format: StoreID|EntryID|InternetMessageId
        /// </summary>
        private string GetEmailReference(Outlook.MailItem mail)
        {
            if (mail == null)
                return string.Empty;

            try
            {
                // Get StoreID from the mail's parent Store
                string storeId = mail.Parent is Outlook.Folder folder
                    ? folder.Store.StoreID
                    : string.Empty;

                // Get EntryID from the mail item
                string entryId = mail.EntryID ?? string.Empty;

                // Get InternetMessageId using PropertyAccessor
                const string PR_INTERNET_MESSAGE_ID =
                    "http://schemas.microsoft.com/mapi/proptag/0x1035001E";

                string internetMessageId = string.Empty;
                try
                {
                    var accessor = mail.PropertyAccessor;
                    object value = accessor.GetProperty(PR_INTERNET_MESSAGE_ID);
                    internetMessageId = value?.ToString() ?? string.Empty;
                }
                catch
                {
                    internetMessageId = string.Empty;
                }

                // Validate all three components exist
                if (string.IsNullOrWhiteSpace(storeId) ||
                    string.IsNullOrWhiteSpace(entryId) ||
                    string.IsNullOrWhiteSpace(internetMessageId))
                {
                    System.Diagnostics.Debug.WriteLine("Warning: Could not extract complete email reference");
                    return string.Empty;
                }

                // Return pipe-delimited format
                return $"{storeId}|{entryId}|{internetMessageId}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting email reference: {ex.Message}");
                return string.Empty;
            }
        }

        private DateTime GetEmailDate(Outlook.MailItem mail)
        {
            try
            {
                // For sent items, use SentOn. For received items, use ReceivedTime
                var folder = mail.Parent as Outlook.MAPIFolder;
                bool isSentFolder = folder != null &&
                                   (folder.Name.Equals("Sent Items", StringComparison.OrdinalIgnoreCase) ||
                                    folder.Name.Equals("Sent", StringComparison.OrdinalIgnoreCase));

                return isSentFolder ? mail.SentOn : mail.ReceivedTime;
            }
            catch
            {
                // Fallback to ReceivedTime if detection fails
                return mail.ReceivedTime;
            }
        }

        private int GetDefaultDueDays()
        {
            int days = Properties.Settings.Default.DefaultDueDays;
            if (days > 0)
                return days;

            return 7; // Fallback default
        }

        private async void ProcessQueue()
        {
            if (_isProcessingQueue)
            {
                SetStatus($"Queued ({_actionQueue.Count} pending)");
                return;
            }

            _isProcessingQueue = true;

            // Disable linked action editing during processing
            Dispatcher.Invoke(() =>
            {
                if (LinkedActionBorder != null)
                    LinkedActionBorder.IsEnabled = false;
            });

            while (_actionQueue.Count > 0)
            {
                var action = _actionQueue.Dequeue();
                try
                {
                    await action();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Queue processing error: {ex.Message}");
                    SetStatus($"Error: {ex.Message}");
                }

                // Small delay between processing
                await Task.Delay(500);
            }

            _isProcessingQueue = false;

            // Re-enable linked action editing
            Dispatcher.Invoke(() =>
            {
                var selectedItem = ActionComboBox.SelectedItem as ActionDropdownItem;
                if (LinkedActionBorder != null && selectedItem?.ItemType == ActionDropdownItemType.Action)
                    LinkedActionBorder.IsEnabled = true;
            });

            SetStatus("Ready");
        }

        private void SetStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text = message;
                System.Diagnostics.Debug.WriteLine($"Status: {message}");
            });
        }

        private async Task LoadRelatedMessagesAsync(ActionItem action)
        {
            if (action == null)
            {
                RelatedMessagesListView.ItemsSource = null;
                return;
            }

            try
            {
                // Parse email references from action
                var emailReferences = action.ParseActiveMessageIds();

                if (emailReferences.Count == 0)
                {
                    RelatedMessagesListView.ItemsSource = null;
                    return;
                }

                SetStatus($"Loading {emailReferences.Count} related messages...");

                // Retrieve emails in background thread
                var relatedEmails = await Task.Run(() =>
                    _emailRetrievalService.RetrieveEmailsByReferences(emailReferences));

                // Update UI on UI thread
                Dispatcher.Invoke(() =>
                {
                    RelatedMessagesListView.ItemsSource = relatedEmails;
                    SetStatus($"Loaded {relatedEmails.Count} related messages");
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading related messages: {ex.Message}");
                SetStatus("Error loading related messages");
                Dispatcher.Invoke(() =>
                {
                    RelatedMessagesListView.ItemsSource = null;
                });
            }
        }

        private void RelatedMessagesListView_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var selectedItem = RelatedMessagesListView.SelectedItem as RelatedEmailItem;
            if (selectedItem?.MailItem != null)
            {
                try
                {
                    // Open email in modeless window (non-blocking)
                    selectedItem.MailItem.Display(false);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error opening email: {ex.Message}");
                    MessageBox.Show($"Error opening email:\n{ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void OpenAllMessagesButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (RelatedMessagesListView.ItemsSource == null)
                    return;

                var relatedEmails = RelatedMessagesListView.ItemsSource as System.Collections.Generic.List<RelatedEmailItem>;
                if (relatedEmails == null || relatedEmails.Count == 0)
                {
                    MessageBox.Show("No related messages to open.", "Information",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                int opened = 0;
                int failed = 0;

                foreach (var email in relatedEmails)
                {
                    if (email?.MailItem != null)
                    {
                        try
                        {
                            // Open email in modeless window (non-blocking)
                            email.MailItem.Display(false);
                            opened++;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error opening email: {ex.Message}");
                            failed++;
                        }
                    }
                }

                if (failed > 0)
                {
                    MessageBox.Show($"Opened {opened} messages. Failed to open {failed} messages.", "Open All",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    SetStatus($"Opened {opened} related messages");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error opening all messages: {ex.Message}");
                MessageBox.Show($"Error opening messages:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #region Compose Mode Detection

        /// <summary>
        /// Called when a compose window (Inspector) is opened OR when inline compose is detected
        /// </summary>
        public void OnComposeItemActivated(Outlook.MailItem mail, Outlook.Inspector inspector)
        {
            System.Diagnostics.Debug.WriteLine($"OnComposeItemActivated: {mail.Subject ?? "(No Subject)"}");
            System.Diagnostics.Debug.WriteLine($"  Inspector: {(inspector != null ? "Popup window" : "Inline editing")}");

            _composeMail = mail;
            _composeInspector = inspector; // Can be null for inline editing
            _isComposeMode = true;

            // Load any existing deferred data
            _currentDeferredData = LoadDeferredData(mail);

            // Hook inspector close event ONLY if popup window (inspector is not null)
            if (inspector != null)
            {
                try
                {
                    ((Outlook.InspectorEvents_10_Event)inspector).Close += Inspector_Close;
                    System.Diagnostics.Debug.WriteLine("  Hooked Inspector.Close event for popup window");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"  Error hooking Inspector.Close: {ex.Message}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("  Inline compose - no Inspector to hook (will exit via Explorer_SelectionChange)");
            }

            // Update UI for compose mode
            UpdateUIForComposeMode();

            System.Diagnostics.Debug.WriteLine($"  Compose mode activated. Deferred data: {(_currentDeferredData != null ? _currentDeferredData.Mode : "None")}");
        }

        /// <summary>
        /// Called when the compose window is closed
        /// </summary>
        private void Inspector_Close()
        {
            System.Diagnostics.Debug.WriteLine("Inspector_Close event fired");
            OnComposeItemDeactivated();
        }

        /// <summary>
        /// Deactivates compose mode and restores normal UI
        /// </summary>
        public void OnComposeItemDeactivated()
        {
            System.Diagnostics.Debug.WriteLine("OnComposeItemDeactivated");

            // Unhook inspector event
            if (_composeInspector != null)
            {
                try
                {
                    ((Outlook.InspectorEvents_10_Event)_composeInspector).Close -= Inspector_Close;
                    System.Diagnostics.Debug.WriteLine("  Unhooked Inspector.Close event");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"  Error unhooking Inspector.Close: {ex.Message}");
                }
            }

            _composeMail = null;
            _composeInspector = null;
            _isComposeMode = false;
            _currentDeferredData = null;

            // Restore normal UI
            Dispatcher.Invoke(() =>
            {
                UpdateButtonVisualsForSchedule(); // Reset button visuals
                UpdateButtonStates();
            });

            System.Diagnostics.Debug.WriteLine("  Compose mode deactivated");
        }

        /// <summary>
        /// Updates the UI to reflect compose mode vs normal mode
        /// </summary>
        private void UpdateUIForComposeMode()
        {
            Dispatcher.Invoke(() =>
            {
                // Update button visuals based on scheduled state
                UpdateButtonVisualsForSchedule();

                // Update button enabled states
                UpdateButtonStates();
            });
        }

        #endregion

        #region Deferred Action - UserProperties Helpers

        /// <summary>
        /// Saves deferred action data to the mail item's UserProperties
        /// </summary>
        private void SaveDeferredData(Outlook.MailItem mail, DeferredActionData data)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== SaveDeferredData START ===");
                System.Diagnostics.Debug.WriteLine($"Mail subject: {mail.Subject}");
                System.Diagnostics.Debug.WriteLine($"Data Mode: {data.Mode}, ActionID: {data.ActionID}");

                var json = JsonSerializer.Serialize(data);
                System.Diagnostics.Debug.WriteLine($"Serialized JSON: {json}");

                var props = mail.UserProperties;
                System.Diagnostics.Debug.WriteLine($"Current UserProperties count: {props.Count}");

                var prop = props.Find(DEFERRED_PROPERTY_NAME);

                if (prop == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Creating new property: {DEFERRED_PROPERTY_NAME}");
                    prop = props.Add(DEFERRED_PROPERTY_NAME, Outlook.OlUserPropertyType.olText);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Updating existing property: {DEFERRED_PROPERTY_NAME}");
                }

                prop.Value = json;
                mail.Save();

                System.Diagnostics.Debug.WriteLine($"✓ Property saved and mail.Save() called");
                System.Diagnostics.Debug.WriteLine($"=== SaveDeferredData END ===");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR saving deferred data: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Loads deferred action data from the mail item's UserProperties
        /// </summary>
        private DeferredActionData LoadDeferredData(Outlook.MailItem mail)
        {
            try
            {
                var props = mail.UserProperties;
                var prop = props.Find(DEFERRED_PROPERTY_NAME);

                if (prop == null || string.IsNullOrEmpty(prop.Value?.ToString()))
                {
                    return null;
                }

                var json = prop.Value.ToString();
                var data = JsonSerializer.Deserialize<DeferredActionData>(json);

                System.Diagnostics.Debug.WriteLine($"Loaded deferred data: {json}");
                return data;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading deferred data: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Clears deferred action data from the mail item's UserProperties
        /// </summary>
        private void ClearDeferredData(Outlook.MailItem mail)
        {
            try
            {
                var props = mail.UserProperties;
                var prop = props.Find(DEFERRED_PROPERTY_NAME);

                if (prop != null)
                {
                    prop.Delete();

                    // Remove "Tracked" category when toggle is turned off (fire-and-forget)
                    _emailCategoryService.RemoveTrackedCategory(mail);

                    mail.Save();
                    System.Diagnostics.Debug.WriteLine("Cleared deferred data");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing deferred data: {ex.Message}");
            }
        }

        #endregion

        #region Deferred Action Execution

        /// <summary>
        /// Executes a deferred create action when the email is sent
        /// </summary>
        public async Task ExecuteDeferredCreateAsync(Outlook.MailItem sentMail)
        {
            System.Diagnostics.Debug.WriteLine($"=== ExecuteDeferredCreateAsync: {sentMail.Subject} ===");

            try
            {
                // Get final email data from the sent item
                string body = sentMail.Body ?? "";
                string senderEmail = sentMail.SenderEmailAddress ?? "";
                string toRecipients = sentMail.To ?? "";
                string subject = sentMail.Subject ?? "";

                System.Diagnostics.Debug.WriteLine("Running auto-classification...");

                // AUTO-CLASSIFY Project and Package
                var classification = await Task.Run(() =>
                    _classifierService.Classify(subject, body, senderEmail, toRecipients));

                string project = classification.SuggestedProjectID ?? "Random";
                string package = classification.SuggestedPackageID ?? "";

                System.Diagnostics.Debug.WriteLine($"Classification: Project={project}, Package={package}");
                System.Diagnostics.Debug.WriteLine("Running LLM extraction...");

                // Get LLM suggestions for Title, BallHolder, Description
                LLMExtractionResult extraction;
                try
                {
                    extraction = await _llmService.GetExtractionAsync(body, senderEmail, subject);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LLM extraction failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsLLMError(sentMail);
                    throw;
                }

                System.Diagnostics.Debug.WriteLine($"LLM extraction: Title={extraction.Title}, BallHolder={extraction.BallHolder}");

                string title = extraction.Title;
                string ballHolder = extraction.BallHolder;
                string description = extraction.Description;

                // Get email metadata
                string conversationId = sentMail.ConversationID;
                string emailReference = GetEmailReference(sentMail);
                DateTime emailDate = GetEmailDate(sentMail);
                int dueDays = GetDefaultDueDays();

                System.Diagnostics.Debug.WriteLine("Creating action in Google Sheets...");

                // Create the action (no confirmation dialog in deferred mode)
                try
                {
                    await _googleSheetsService.CreateActionAsync(
                        project,
                        package,
                        title,
                        ballHolder,
                        conversationId,
                        emailReference,
                        description,
                        emailDate,
                        dueDays
                    );

                    // Tag email as tracked (fire-and-forget)
                    _emailCategoryService.MarkEmailAsTracked(sentMail);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Google Sheets CreateAction (deferred) failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsSheetsError(sentMail);
                    throw;
                }

                System.Diagnostics.Debug.WriteLine($"=== Deferred create SUCCESS: {title} ===");

                // Update status
                SetStatus($"✓ Action created: {title}");

                // Reload actions
                await LoadActionsAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"=== Deferred create FAILED: {ex.Message} ===");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                SetStatus($"✗ Failed to create action: {ex.Message}");
                throw; // Re-throw so ThisAddIn can handle the error
            }
        }

        /// <summary>
        /// Executes a deferred update action when the email is sent
        /// </summary>
        public async Task ExecuteDeferredUpdateAsync(Outlook.MailItem sentMail, int actionId)
        {
            System.Diagnostics.Debug.WriteLine($"=== ExecuteDeferredUpdateAsync: ActionID={actionId}, Email={sentMail.Subject} ===");

            try
            {
                // Find the action by ID
                var selectedAction = _openActions?.FirstOrDefault(a => a.Id == actionId);

                if (selectedAction == null)
                {
                    System.Diagnostics.Debug.WriteLine($"WARNING: Action ID {actionId} not found in open actions list");
                    // Try to reload actions and search again
                    await LoadActionsAsync();
                    selectedAction = _openActions?.FirstOrDefault(a => a.Id == actionId);

                    if (selectedAction == null)
                    {
                        throw new Exception($"Action ID {actionId} not found");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Found action: {selectedAction.Title}");

                // Get final email body from sent item
                string body = sentMail.Body ?? "";
                string currentContext = selectedAction.HistoryLog ?? "";
                string currentBallHolder = selectedAction.BallHolder ?? "";

                System.Diagnostics.Debug.WriteLine("Running LLM delta analysis...");

                // Get LLM delta
                LLMDeltaResult delta;
                try
                {
                    delta = await _llmService.GetDeltaAsync(body, currentContext, currentBallHolder);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LLM delta failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsLLMError(sentMail);
                    throw;
                }

                System.Diagnostics.Debug.WriteLine($"LLM delta: BallHolder={delta.NewBallHolder}, Summary={delta.UpdateSummary?.Substring(0, Math.Min(50, delta.UpdateSummary?.Length ?? 0))}");

                string ballHolder = delta.NewBallHolder;
                string updateNote = delta.UpdateSummary;

                // Get email metadata
                string emailReference = GetEmailReference(sentMail);
                DateTime emailDate = GetEmailDate(sentMail);
                int dueDays = GetDefaultDueDays();

                System.Diagnostics.Debug.WriteLine("Updating action in Google Sheets...");

                // Update the action (no confirmation dialog in deferred mode)
                try
                {
                    await _googleSheetsService.UpdateActionAsync(
                        selectedAction.Id,
                        emailReference,
                        ballHolder,
                        updateNote,
                        emailDate,
                        dueDays
                    );

                    // Tag email as tracked (fire-and-forget)
                    _emailCategoryService.MarkEmailAsTracked(sentMail);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Google Sheets UpdateAction (deferred) failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsSheetsError(sentMail);
                    throw;
                }

                System.Diagnostics.Debug.WriteLine($"=== Deferred update SUCCESS: {selectedAction.Title} ===");

                // Update status
                SetStatus($"✓ Action updated: {selectedAction.Title}");

                // Reload actions
                await LoadActionsAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"=== Deferred update FAILED: {ex.Message} ===");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                SetStatus($"✗ Failed to update action: {ex.Message}");
                throw; // Re-throw so ThisAddIn can handle the error
            }
        }

        /// <summary>
        /// Executes a deferred "create multiple" action when the email is sent
        /// </summary>
        public async Task ExecuteDeferredCreateMultipleAsync(Outlook.MailItem sentMail)
        {
            System.Diagnostics.Debug.WriteLine($"=== ExecuteDeferredCreateMultipleAsync: {sentMail.Subject} ===");

            try
            {
                // Just call the regular CreateMultiple processing logic
                await ProcessCreateMultipleActionsAsync(sentMail);

                System.Diagnostics.Debug.WriteLine("=== Deferred create multiple SUCCESS ===");
                SetStatus("✓ Multiple actions created");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"=== Deferred create multiple FAILED: {ex.Message} ===");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                SetStatus($"✗ Failed to create multiple actions: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Executes a deferred close action when the email is sent
        /// </summary>
        public async Task ExecuteDeferredCloseAsync(Outlook.MailItem sentMail, int actionId)
        {
            System.Diagnostics.Debug.WriteLine($"=== ExecuteDeferredCloseAsync: ActionID={actionId}, Email={sentMail.Subject} ===");

            try
            {
                // Find the action by ID
                var selectedAction = _openActions?.FirstOrDefault(a => a.Id == actionId);

                if (selectedAction == null)
                {
                    System.Diagnostics.Debug.WriteLine($"WARNING: Action ID {actionId} not found in open actions list");
                    await LoadActionsAsync();
                    selectedAction = _openActions?.FirstOrDefault(a => a.Id == actionId);

                    if (selectedAction == null)
                    {
                        throw new Exception($"Action ID {actionId} not found");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Found action: {selectedAction.Title}");

                // Get email details
                DateTime emailDate = GetEmailDate(sentMail);
                string closingEmailReference = GetEmailReference(sentMail);

                if (string.IsNullOrWhiteSpace(closingEmailReference))
                {
                    throw new Exception("Could not extract email reference");
                }

                // Generate LLM closure summary
                string actionContext = $"Title: {selectedAction.Title}\nProject: {selectedAction.Project}\nPackage: {selectedAction.Package}";
                string closureNote;

                try
                {
                    System.Diagnostics.Debug.WriteLine("Analyzing email for closure summary...");
                    closureNote = await _llmService.GetClosureSummaryAsync(sentMail.Body, actionContext);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LLM closure summary failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsLLMError(sentMail);
                    closureNote = "Action closed";
                }

                System.Diagnostics.Debug.WriteLine("Closing action in Google Sheets...");

                // Close the action
                try
                {
                    await _googleSheetsService.CloseActionAsync(
                        selectedAction.Id,
                        closureNote,
                        closingEmailReference,
                        emailDate
                    );

                    // Tag email as tracked (fire-and-forget)
                    _emailCategoryService.MarkEmailAsTracked(sentMail);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Google Sheets CloseAction (deferred) failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsSheetsError(sentMail);
                    throw;
                }

                System.Diagnostics.Debug.WriteLine($"=== Deferred close SUCCESS: {selectedAction.Title} ===");
                SetStatus($"✓ Action closed: {selectedAction.Title}");

                // Reload actions
                await LoadActionsAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"=== Deferred close FAILED: {ex.Message} ===");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                SetStatus($"✗ Failed to close action: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Executes a deferred reopen action when the email is sent
        /// </summary>
        public async Task ExecuteDeferredReopenAsync(Outlook.MailItem sentMail, int actionId)
        {
            System.Diagnostics.Debug.WriteLine($"=== ExecuteDeferredReopenAsync: ActionID={actionId}, Email={sentMail.Subject} ===");

            try
            {
                // Find the action by ID
                var selectedAction = _openActions?.FirstOrDefault(a => a.Id == actionId);

                if (selectedAction == null)
                {
                    System.Diagnostics.Debug.WriteLine($"WARNING: Action ID {actionId} not found");
                    await LoadActionsAsync();
                    selectedAction = _openActions?.FirstOrDefault(a => a.Id == actionId);

                    if (selectedAction == null)
                    {
                        throw new Exception($"Action ID {actionId} not found");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Found action: {selectedAction.Title}");

                DateTime emailDate = GetEmailDate(sentMail);
                int dueDays = GetDefaultDueDays();

                System.Diagnostics.Debug.WriteLine("Reopening action in Google Sheets...");

                try
                {
                    await _googleSheetsService.SetStatusAsync(selectedAction.Id, "Open", emailDate, dueDays);

                    // Tag email as tracked (fire-and-forget)
                    _emailCategoryService.MarkEmailAsTracked(sentMail);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Google Sheets SetStatus (deferred) failed: {ex.Message}");
                    _emailCategoryService.MarkEmailAsSheetsError(sentMail);
                    throw;
                }

                System.Diagnostics.Debug.WriteLine($"=== Deferred reopen SUCCESS: {selectedAction.Title} ===");
                SetStatus($"✓ Action reopened: {selectedAction.Title}");

                // Reload actions
                await LoadActionsAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"=== Deferred reopen FAILED: {ex.Message} ===");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                SetStatus($"✗ Failed to reopen action: {ex.Message}");
                throw;
            }
        }

        #endregion

        #region Command Center Dashboard

        private void OverdueActionsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedAction = OverdueActionsComboBox.SelectedItem as ActionItem;
            if (selectedAction != null)
            {
                WithMeActionsComboBox.SelectedItem = null;
                LoadActionIntoDashboard(selectedAction);
            }
        }

        private void WithMeActionsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedAction = WithMeActionsComboBox.SelectedItem as ActionItem;
            if (selectedAction != null)
            {
                OverdueActionsComboBox.SelectedItem = null;
                LoadActionIntoDashboard(selectedAction);
            }
        }

        private void ExpandUserFilterButton_Click(object sender, RoutedEventArgs e)
        {
            _isUserFilterExpanded = true;

            // Get all unique users from BallHolder fields (split by semicolon)
            var allUsers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (_openActions != null)
            {
                foreach (var action in _openActions)
                {
                    if (!string.IsNullOrWhiteSpace(action.BallHolder))
                    {
                        // Split by semicolon and add each user
                        var users = action.BallHolder.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var user in users)
                        {
                            var trimmedUser = user.Trim();
                            if (!string.IsNullOrWhiteSpace(trimmedUser))
                            {
                                allUsers.Add(trimmedUser);
                            }
                        }
                    }
                }
            }

            // Populate user filter dropdown
            var sortedUsers = allUsers.OrderBy(u => u).ToList();
            UserFilterComboBox.ItemsSource = sortedUsers;

            // Select current user by default if they exist in the list
            var currentUserMatch = sortedUsers.FirstOrDefault(u =>
                u.Equals(_currentUserName, StringComparison.OrdinalIgnoreCase) ||
                IsAssignedToUser(u, _currentUserName));

            if (currentUserMatch != null)
            {
                UserFilterComboBox.SelectedItem = currentUserMatch;
            }
            else if (sortedUsers.Count > 0)
            {
                UserFilterComboBox.SelectedIndex = 0;
            }

            // Update UI
            WithMeLabel.Text = "With...";
            ExpandUserFilterButton.Visibility = Visibility.Collapsed;
            UserFilterComboBox.Visibility = Visibility.Visible;
            ResetToMeButton.Visibility = Visibility.Visible;
        }

        private void ResetToMeButton_Click(object sender, RoutedEventArgs e)
        {
            _isUserFilterExpanded = false;
            _selectedFilterUser = null;

            // Update UI
            ExpandUserFilterButton.Visibility = Visibility.Visible;
            UserFilterComboBox.Visibility = Visibility.Collapsed;
            ResetToMeButton.Visibility = Visibility.Collapsed;

            // Refresh dashboard to show "With Me" again
            RefreshDashboard();
        }

        private void UserFilterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (UserFilterComboBox.SelectedItem is string selectedUser)
            {
                _selectedFilterUser = selectedUser;

                // Filter actions where selected user appears in BallHolder
                var filteredActions = _openActions
                    .Where(a =>
                    {
                        string status = a.Status ?? "";
                        bool isOpen = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase);

                        bool isOverdue = isOpen
                            && a.DueDate.HasValue
                            && a.DueDate.Value.Date < DateTime.Today;

                        bool hasUser = !string.IsNullOrWhiteSpace(a.BallHolder) &&
                                      IsUserInBallHolder(a.BallHolder, selectedUser);

                        return isOpen && hasUser && !isOverdue;
                    })
                    .OrderBy(a => a.DueDate)
                    .ToList();

                // Update the WithMe section
                _withMeActions = filteredActions;
                WithMeActionsComboBox.ItemsSource = _withMeActions;
                WithMeLabel.Text = $"With {GetShortName(selectedUser)}: {_withMeActions.Count}";

                // Update visibility
                WithMeSection.Visibility = _withMeActions.Count > 0
                    ? Visibility.Visible
                    : Visibility.Collapsed;

                // Update collapsed summary if Open Actions is collapsed
                // Always show "With me:" based on current user's count
                if (_isOpenActionsCollapsed)
                {
                    int myActionsCount = _openActions
                        .Where(a =>
                        {
                            string status = a.Status ?? "";
                            bool isOverdue = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                                && a.DueDate.HasValue
                                && a.DueDate.Value.Date < DateTime.Today;

                            bool isWithMe = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                                && !string.IsNullOrWhiteSpace(a.BallHolder)
                                && IsAssignedToUser(a.BallHolder, _currentUserName);

                            return isWithMe && !isOverdue;
                        })
                        .Count();

                    OpenActionsCollapsedSummary.Text = $"Overdue: {_overdueActions?.Count ?? 0}, With me: {myActionsCount}";
                }
            }
        }

        /// <summary>
        /// Checks if a user appears in the BallHolder field (handles multi-user assignments with semicolons)
        /// </summary>
        private bool IsUserInBallHolder(string ballHolder, string userName)
        {
            if (string.IsNullOrWhiteSpace(ballHolder) || string.IsNullOrWhiteSpace(userName))
                return false;

            // Split by semicolon to handle multiple users
            var users = ballHolder.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var user in users)
            {
                var trimmedUser = user.Trim();
                // Check if this user matches the target
                if (trimmedUser.Equals(userName, StringComparison.OrdinalIgnoreCase) ||
                    IsAssignedToUser(trimmedUser, userName))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Gets a shortened version of a user name for display (first name only)
        /// </summary>
        private string GetShortName(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
                return "";

            // If name contains a space, return first part
            var parts = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length > 0 ? parts[0] : fullName;
        }

        private void SearchActionButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string searchText = ActionIdSearchBox.Text?.Trim();

                if (string.IsNullOrWhiteSpace(searchText) || searchText == "Action ID")
                {
                    MessageBox.Show("Please enter an Action ID to search.", "No Input",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (!int.TryParse(searchText, out int actionId))
                {
                    MessageBox.Show("Action ID must be a number.", "Invalid Input",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var matchingActions = _openActions?.Where(a => a.Id == actionId).ToList();

                if (matchingActions == null || matchingActions.Count == 0)
                {
                    MessageBox.Show($"No action found with ID {actionId}.", "Not Found",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Clear ComboBox selections since this is a manual search
                OverdueActionsComboBox.SelectedItem = null;
                WithMeActionsComboBox.SelectedItem = null;

                if (matchingActions.Count > 1)
                {
                    var resolvedAction = ShowDuplicateActionResolver(matchingActions);
                    if (resolvedAction != null)
                    {
                        LoadActionIntoDashboard(resolvedAction);
                    }
                }
                else
                {
                    LoadActionIntoDashboard(matchingActions[0]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error searching for action:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ActionIdSearchBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                SearchActionButton_Click(sender, e);
            }
        }

        private void ActionIdSearchBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (ActionIdSearchBox.Text == "Action ID")
            {
                ActionIdSearchBox.Text = "";
                ActionIdSearchBox.Foreground = (SolidColorBrush)FindResource("TextPrimary");
            }
        }

        private void ActionIdSearchBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ActionIdSearchBox.Text))
            {
                ActionIdSearchBox.Text = "Action ID";
                ActionIdSearchBox.Foreground = (SolidColorBrush)FindResource("TextSecondary");
            }
        }

        private async void LoadActionIntoDashboard(ActionItem action)
        {
            if (action == null) return;

            try
            {
                _dashboardSelectedAction = action;

                // Populate action details
                DashboardProjectText.Text = action.Project ?? "";
                DashboardPackageText.Text = action.Package ?? "";
                DashboardTitleText.Text = action.Title ?? "";
                DashboardBallHolderText.Text = action.BallHolder ?? "";

                // Show results section
                DashboardResultsSection.Visibility = Visibility.Visible;

                // Load related emails
                var emailReferences = action.ParseActiveMessageIds();

                if (emailReferences.Count == 0)
                {
                    DashboardRelatedMessagesListView.ItemsSource = null;
                    return;
                }

                SetStatus($"Loading {emailReferences.Count} related messages for Action {action.Id}...");

                var relatedEmails = await Task.Run(() =>
                    _emailRetrievalService.RetrieveEmailsByReferences(emailReferences));

                Dispatcher.Invoke(() =>
                {
                    DashboardRelatedMessagesListView.ItemsSource = relatedEmails;
                    SetStatus($"Dashboard: Loaded {relatedEmails.Count} messages for Action {action.Id}");
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading action into dashboard: {ex.Message}");
                SetStatus("Error loading action details");
            }
        }

        private ActionItem ShowDuplicateActionResolver(List<ActionItem> duplicates)
        {
            var dialog = new Window
            {
                Title = "Multiple Actions Found",
                Width = 450,
                Height = 300,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize
            };

            var stackPanel = new StackPanel { Margin = new Thickness(16) };

            stackPanel.Children.Add(new TextBlock
            {
                Text = $"Found {duplicates.Count} actions with this ID. Please select one:",
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap
            });

            var listView = new ListView
            {
                Height = 180,
                ItemsSource = duplicates
            };

            listView.View = new GridView();
            ((GridView)listView.View).Columns.Add(new GridViewColumn
            {
                Header = "Project",
                DisplayMemberBinding = new System.Windows.Data.Binding("Project"),
                Width = 80
            });
            ((GridView)listView.View).Columns.Add(new GridViewColumn
            {
                Header = "Title",
                DisplayMemberBinding = new System.Windows.Data.Binding("Title"),
                Width = 200
            });
            ((GridView)listView.View).Columns.Add(new GridViewColumn
            {
                Header = "Ball Holder",
                DisplayMemberBinding = new System.Windows.Data.Binding("BallHolder"),
                Width = 100
            });

            stackPanel.Children.Add(listView);

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };

            var selectButton = new Button
            {
                Content = "Select",
                Width = 80,
                Margin = new Thickness(0, 0, 8, 0),
                Style = (Style)FindResource("ModernButtonPrimary")
            };

            var cancelButton = new Button
            {
                Content = "Cancel",
                Width = 80,
                Style = (Style)FindResource("ModernButtonSecondary")
            };

            ActionItem selectedAction = null;

            selectButton.Click += (s, e) =>
            {
                if (listView.SelectedItem != null)
                {
                    selectedAction = listView.SelectedItem as ActionItem;
                    dialog.DialogResult = true;
                    dialog.Close();
                }
            };

            cancelButton.Click += (s, e) =>
            {
                dialog.DialogResult = false;
                dialog.Close();
            };

            buttonPanel.Children.Add(selectButton);
            buttonPanel.Children.Add(cancelButton);
            stackPanel.Children.Add(buttonPanel);

            dialog.Content = stackPanel;
            dialog.ShowDialog();

            return selectedAction;
        }

        private void DashboardRelatedMessagesListView_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var selectedItem = DashboardRelatedMessagesListView.SelectedItem as RelatedEmailItem;
            if (selectedItem?.MailItem != null)
            {
                try
                {
                    selectedItem.MailItem.Display(false);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error opening email: {ex.Message}");
                    MessageBox.Show($"Error opening email:\n{ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DashboardOpenAllButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (DashboardRelatedMessagesListView.ItemsSource == null)
                    return;

                var relatedEmails = DashboardRelatedMessagesListView.ItemsSource
                    as System.Collections.Generic.List<RelatedEmailItem>;

                if (relatedEmails == null || relatedEmails.Count == 0)
                {
                    MessageBox.Show("No related messages to open.", "Information",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                int opened = 0;
                int failed = 0;

                foreach (var email in relatedEmails)
                {
                    try
                    {
                        if (email.MailItem != null)
                        {
                            email.MailItem.Display(false);
                            opened++;
                        }
                    }
                    catch
                    {
                        failed++;
                    }
                }

                SetStatus($"Opened {opened} messages" + (failed > 0 ? $" ({failed} failed)" : ""));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening messages:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Collapse/Expand Handlers

        private void OpenActionsHeader_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // If Linked Action is collapsed, expand it first
            if (_isLinkedActionCollapsed && !_isOpenActionsCollapsed)
            {
                ToggleLinkedActionCollapse();
            }

            ToggleOpenActionsCollapse();
        }

        private void LinkedActionHeader_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // If Open Actions is collapsed, expand it first
            if (_isOpenActionsCollapsed && !_isLinkedActionCollapsed)
            {
                ToggleOpenActionsCollapse();
            }

            ToggleLinkedActionCollapse();
        }

        private void ToggleOpenActionsCollapse()
        {
            _isOpenActionsCollapsed = !_isOpenActionsCollapsed;

            if (_isOpenActionsCollapsed)
            {
                // Save current height before collapsing
                _savedOpenActionsHeight = OpenActionsRow.Height;

                // Collapse: hide content, show summary, and resize to minimal height
                OpenActionsContent.Visibility = Visibility.Collapsed;
                OpenActionsExpandIcon.Text = "▶";

                // Show collapsed summary - always show current user's count, not filtered user
                int overdueCount = _overdueActions?.Count ?? 0;
                int myActionsCount = _openActions?
                    .Where(a =>
                    {
                        string status = a.Status ?? "";
                        bool isOverdue = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                            && a.DueDate.HasValue
                            && a.DueDate.Value.Date < DateTime.Today;

                        bool isWithMe = !status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
                            && !string.IsNullOrWhiteSpace(a.BallHolder)
                            && IsAssignedToUser(a.BallHolder, _currentUserName);

                        return isWithMe && !isOverdue;
                    })
                    .Count() ?? 0;
                OpenActionsCollapsedSummary.Text = $"Overdue: {overdueCount}, With me: {myActionsCount}";
                OpenActionsCollapsedSummary.Visibility = Visibility.Visible;

                // Resize to minimal height (just header)
                OpenActionsRow.MinHeight = 0;
                OpenActionsRow.Height = new GridLength(50);
            }
            else
            {
                // Expand: show content, hide summary, restore previous height
                OpenActionsContent.Visibility = Visibility.Visible;
                OpenActionsExpandIcon.Text = "▼";
                OpenActionsCollapsedSummary.Visibility = Visibility.Collapsed;

                // Restore MinHeight
                OpenActionsRow.MinHeight = 100;

                // Restore previous height (or default if not set)
                if (_savedOpenActionsHeight.Value > 0)
                {
                    OpenActionsRow.Height = _savedOpenActionsHeight;
                }
                else
                {
                    OpenActionsRow.Height = new GridLength(1, GridUnitType.Star);
                }
            }
        }

        private void ToggleLinkedActionCollapse()
        {
            _isLinkedActionCollapsed = !_isLinkedActionCollapsed;

            if (_isLinkedActionCollapsed)
            {
                // Save current height before collapsing
                _savedLinkedActionHeight = LinkedActionRow.Height;

                // Collapse: hide the details content and resize to minimal height
                LinkedActionContent.Visibility = Visibility.Collapsed;
                LinkedActionExpandIcon.Text = "▶";

                // Resize to minimal height (just header + dropdown)
                LinkedActionRow.MinHeight = 0;
                LinkedActionRow.Height = new GridLength(92);
            }
            else
            {
                // Expand: show content and restore previous height
                LinkedActionContent.Visibility = Visibility.Visible;
                LinkedActionExpandIcon.Text = "▼";

                // Restore MinHeight
                LinkedActionRow.MinHeight = 150;

                // Restore previous height (or default if not set)
                if (_savedLinkedActionHeight.Value > 0)
                {
                    LinkedActionRow.Height = _savedLinkedActionHeight;
                }
                else
                {
                    LinkedActionRow.Height = new GridLength(1.5, GridUnitType.Star);
                }
            }
        }

        #endregion
    }
}
