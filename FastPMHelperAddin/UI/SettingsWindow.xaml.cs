using System;
using System.Windows;

namespace FastPMHelperAddin.UI
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            try
            {
                // Load UserName
                MyNameInput.Text = Properties.Settings.Default.UserName ?? "Wally Cloud";

                // Load DefaultDueDays (default to 7 if not set)
                int dueDays = Properties.Settings.Default.DefaultDueDays;
                if (dueDays <= 0) dueDays = 7;
                DefaultDueDaysTextBox.Text = dueDays.ToString();

                // Load ConfirmActions
                ConfirmActionsCheckBox.IsChecked = Properties.Settings.Default.ConfirmActions;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
                // Set defaults on error
                MyNameInput.Text = "Wally Cloud";
                DefaultDueDaysTextBox.Text = "7";
                ConfirmActionsCheckBox.IsChecked = false;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Save UserName
                string userName = MyNameInput.Text?.Trim() ?? "Wally Cloud";
                if (string.IsNullOrWhiteSpace(userName))
                    userName = "Wally Cloud";
                Properties.Settings.Default.UserName = userName;

                // Save DefaultDueDays
                if (int.TryParse(DefaultDueDaysTextBox.Text, out int dueDays) && dueDays > 0)
                {
                    Properties.Settings.Default.DefaultDueDays = dueDays;
                }
                else
                {
                    Properties.Settings.Default.DefaultDueDays = 7;
                }

                // Save ConfirmActions
                Properties.Settings.Default.ConfirmActions = ConfirmActionsCheckBox.IsChecked == true;

                // Persist settings
                Properties.Settings.Default.Save();

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
