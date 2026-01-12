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

        private async void LoadActionsAsync()
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
            }
            else
            {
                PopulateLinkedActionFields(selectedAction);
                LinkedActionBorder.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204)); // Pastel yellow

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
                    string messageId = GetInternetMessageId(mail);
                    DateTime emailDate = GetEmailDate(mail);
                    int dueDays = GetDefaultDueDays();

                    await _googleSheetsService.CreateActionAsync(
                        project,        // NEW parameter
                        package,        // NEW parameter
                        title,
                        ballHolder,
                        conversationId,
                        messageId,
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

                    string messageId = GetInternetMessageId(mail);
                    DateTime emailDate = GetEmailDate(mail);
                    int dueDays = GetDefaultDueDays();

                    await _googleSheetsService.UpdateActionAsync(
                        selectedAction.Id,
                        messageId,
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
                DateTime emailDate = GetEmailDate(_currentMail);
                int dueDays = GetDefaultDueDays();

                await _googleSheetsService.SetStatusAsync(selectedAction.Id, "Closed", emailDate, dueDays);

                MessageBox.Show("Action closed.", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                LoadActionsAsync();
            }
            catch (Exception ex)
            {
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
    }
}
