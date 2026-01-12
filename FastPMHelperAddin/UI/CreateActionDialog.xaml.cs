using System.Windows;

namespace FastPMHelperAddin.UI
{
    public partial class CreateActionDialog : Window
    {
        public string Project
        {
            get => ProjectTextBox.Text;
            set => ProjectTextBox.Text = value;
        }

        public string Package
        {
            get => PackageTextBox.Text;
            set => PackageTextBox.Text = value;
        }

        public string ActionTitle
        {
            get => TitleTextBox.Text;
            set => TitleTextBox.Text = value;
        }

        public string BallHolder
        {
            get => BallHolderTextBox.Text;
            set => BallHolderTextBox.Text = value;
        }

        public string Description
        {
            get => DescriptionTextBox.Text;
            set => DescriptionTextBox.Text = value;
        }

        public CreateActionDialog()
        {
            InitializeComponent();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ActionTitle))
            {
                MessageBox.Show("Please enter an action title.", "Validation",
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
