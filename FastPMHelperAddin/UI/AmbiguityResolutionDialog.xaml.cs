using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using FastPMHelperAddin.Models;

namespace FastPMHelperAddin.UI
{
    public partial class AmbiguityResolutionDialog : Window
    {
        public string AmbiguityReason { get; set; }
        public List<ClassificationCandidate> Candidates { get; set; }
        public List<ClassificationCandidate> SelectedCandidates { get; private set; }

        private Dictionary<string, RadioButton> _projectRadioButtons;
        private Dictionary<string, RadioButton> _packageRadioButtons;

        public AmbiguityResolutionDialog()
        {
            InitializeComponent();
            _projectRadioButtons = new Dictionary<string, RadioButton>();
            _packageRadioButtons = new Dictionary<string, RadioButton>();

            Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            // Display reason
            ReasonTextBlock.Text = AmbiguityReason ?? "Please select the correct classification:";

            // Populate candidates
            PopulateCandidates();
        }

        private void PopulateCandidates()
        {
            if (Candidates == null || Candidates.Count == 0)
                return;

            // Group by Type
            var projectCandidates = Candidates.Where(c => c.Type == "PROJECT").ToList();
            var packageCandidates = Candidates.Where(c => c.Type == "PACKAGE").ToList();

            // Add PROJECT section
            if (projectCandidates.Any())
            {
                AddSectionHeader("Projects:");
                foreach (var candidate in projectCandidates.OrderByDescending(c => c.Score))
                {
                    AddCandidateRadioButton(candidate, "ProjectGroup", _projectRadioButtons);
                }

                // Add spacing
                CandidatesPanel.Children.Add(new Border { Height = 10 });
            }

            // Add PACKAGE section
            if (packageCandidates.Any())
            {
                AddSectionHeader("Packages:");
                foreach (var candidate in packageCandidates.OrderByDescending(c => c.Score))
                {
                    AddCandidateRadioButton(candidate, "PackageGroup", _packageRadioButtons);
                }
            }
        }

        private void AddSectionHeader(string text)
        {
            var header = new TextBlock
            {
                Text = text,
                FontWeight = FontWeights.Bold,
                FontSize = 14,
                Margin = new Thickness(0, 0, 0, 5)
            };
            CandidatesPanel.Children.Add(header);
        }

        private void AddCandidateRadioButton(ClassificationCandidate candidate, string groupName,
            Dictionary<string, RadioButton> buttonDict)
        {
            var radioButton = new RadioButton
            {
                Content = $"{candidate.Name} (Score: {candidate.Score})",
                GroupName = groupName,
                Margin = new Thickness(10, 3, 0, 3),
                Tag = candidate
            };

            // Select first option by default
            if (buttonDict.Count == 0)
                radioButton.IsChecked = true;

            buttonDict[candidate.Name] = radioButton;
            CandidatesPanel.Children.Add(radioButton);
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedCandidates = new List<ClassificationCandidate>();

            // Get selected project
            var selectedProjectButton = _projectRadioButtons.Values.FirstOrDefault(rb => rb.IsChecked == true);
            if (selectedProjectButton != null)
            {
                SelectedCandidates.Add((ClassificationCandidate)selectedProjectButton.Tag);
            }

            // Get selected package
            var selectedPackageButton = _packageRadioButtons.Values.FirstOrDefault(rb => rb.IsChecked == true);
            if (selectedPackageButton != null)
            {
                SelectedCandidates.Add((ClassificationCandidate)selectedPackageButton.Tag);
            }

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
