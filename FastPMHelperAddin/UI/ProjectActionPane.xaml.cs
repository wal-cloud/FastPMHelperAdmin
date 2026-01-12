using System;
using System.Collections.Generic;
using System.Linq;
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

        private List<ActionItem> _openActions;
        private Outlook.MailItem _currentMail;

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

        public ProjectActionPane()
        {
            InitializeComponent();
            InitializeServices();
            LoadActionsAsync();
        }

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

        // Called from ThisAddIn when email selection changes
        // NOTE: ThisAddIn already wraps this call in Dispatcher.Invoke, so we don't need to do it again
        public void OnEmailSelected(Outlook.MailItem mail)
        {
            System.Diagnostics.Debug.WriteLine($"OnEmailSelected called with: {mail?.Subject ?? "(null)"}");

            _currentMail = mail;

            if (mail == null)
            {
                SelectedEmailSubject.Text = "No email selected";
                ActionComboBox.SelectedItem = null;
                ClearLinkedActionFields();
                UpdateButtonStates();
                return;
            }

            SelectedEmailSubject.Text = mail.Subject ?? "(No Subject)";

            // Auto-match action
            var matchedAction = _matchingService.FindMatchingAction(mail, _openActions);
            ActionComboBox.SelectedItem = matchedAction;

            UpdateButtonStates();

            System.Diagnostics.Debug.WriteLine($"Email selection updated: {mail.Subject}");
        }

        private void ActionComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedAction = ActionComboBox.SelectedItem as ActionItem;

            if (selectedAction == null)
            {
                ClearLinkedActionFields();
                LinkedActionBorder.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Grey
                RelatedMessagesListView.ItemsSource = null;
            }
            else
            {
                PopulateLinkedActionFields(selectedAction);
                LinkedActionBorder.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204)); // Pastel yellow
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
            }

            UpdateButtonStates();
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
                    RefreshActionComboBox();

                    // Re-match current email if one is selected
                    if (_currentMail != null)
                    {
                        var matchedAction = _matchingService.FindMatchingAction(_currentMail, _openActions);
                        ActionComboBox.SelectedItem = matchedAction;
                    }
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
            var selectedAction = ActionComboBox.SelectedItem as ActionItem;
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

                // Refresh ComboBox display
                RefreshActionComboBox();
                ActionComboBox.SelectedItem = selectedAction;

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
            bool hasSelection = ActionComboBox.SelectedItem != null;
            bool hasEmail = _currentMail != null;

            // Enable Create and Create Multiple buttons when email is selected
            CreateMultipleButton.IsEnabled = hasEmail;

            UpdateButton.IsEnabled = hasSelection && hasEmail;
            CloseButton.IsEnabled = hasSelection;
            ReopenButton.IsEnabled = hasSelection;
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
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

        private void CreateMultipleButton_Click(object sender, RoutedEventArgs e)
        {
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
                var extraction = await _llmService.GetExtractionAsync(body, senderEmail, subject);

                bool shouldCreate = true;
                string title = extraction.Title;
                string ballHolder = extraction.BallHolder;
                string description = extraction.Description;

                // Check if confirmation is enabled
                if (ConfirmActionsCheckBox.IsChecked == true)
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
                var extractions = await _llmService.GetMultipleExtractionsAsync(body, senderEmail, subject);

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
                bool confirmActions = ConfirmActionsCheckBox?.IsChecked == true;

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

                        System.Diagnostics.Debug.WriteLine($"Created action {actionNumber}: ID={newActionId}, Title={finalTitle}");
                        successCount++;
                        createdTitles.Add(finalTitle);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error creating action {actionNumber}: {ex.Message}");
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
            var selectedAction = ActionComboBox.SelectedItem as ActionItem;
            if (selectedAction == null || _currentMail == null)
                return;

            // Capture current values before queuing
            var mailToProcess = _currentMail;
            var actionToUpdate = selectedAction;

            // Queue the action
            _actionQueue.Enqueue(() => ProcessUpdateActionAsync(mailToProcess, actionToUpdate));
            ProcessQueue();
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
                if (ConfirmActionsCheckBox.IsChecked == true)
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

                    await _googleSheetsService.UpdateActionAsync(
                        selectedAction.Id,
                        emailReference,
                        ballHolder,
                        updateNote,
                        emailDate,      // SentOn
                        dueDays         // Due days offset
                    );

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
            var selectedAction = ActionComboBox.SelectedItem as ActionItem;
            if (selectedAction == null || _currentMail == null)
                return;

            try
            {
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
                    closureNote = "Action closed";
                }

                SetStatus("Closing action...");

                // Close the action with LLM-generated note, email reference, and date
                await _googleSheetsService.CloseActionAsync(
                    selectedAction.Id,
                    closureNote,
                    closingEmailReference,
                    emailDate
                );

                SetStatus("Action closed successfully");

                await LoadActionsAsync();
            }
            catch (Exception ex)
            {
                SetStatus("Error closing action");
                MessageBox.Show($"Error closing action: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void ReopenButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedAction = ActionComboBox.SelectedItem as ActionItem;
            if (selectedAction == null || _currentMail == null)
                return;

            try
            {
                DateTime emailDate = GetEmailDate(_currentMail);
                int dueDays = GetDefaultDueDays();

                await _googleSheetsService.SetStatusAsync(selectedAction.Id, "Open", emailDate, dueDays);

                MessageBox.Show("Action reopened.", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                LoadActionsAsync();
            }
            catch (Exception ex)
            {
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
            if (int.TryParse(DefaultDueDaysTextBox.Text, out int days) && days > 0)
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
                if (LinkedActionBorder != null && ActionComboBox.SelectedItem != null)
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
    }
}
