using System.Windows;

namespace FastPMHelperAddin.UI
{
    public partial class UpdateActionDialog : Window
    {
        public string BallHolder
        {
            get => BallHolderTextBox.Text;
            set => BallHolderTextBox.Text = value;
        }

        public string UpdateNote
        {
            get => UpdateNoteTextBox.Text;
            set => UpdateNoteTextBox.Text = value;
        }

        public UpdateActionDialog()
        {
            InitializeComponent();
        }

        private void UpdateButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(UpdateNote))
            {
                MessageBox.Show("Please enter an update note.", "Validation",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
