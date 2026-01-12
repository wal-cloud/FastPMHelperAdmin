using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using FastPMHelperAddin.Models;

namespace FastPMHelperAddin.Services
{
    public class AutoClassifierService
    {
        private List<ClassificationRule> _rules;

        public AutoClassifierService()
        {
            _rules = new List<ClassificationRule>();
        }

        /// <summary>
        /// Load rules from Google Sheets API response
        /// </summary>
        public void LoadRules(IList<IList<object>> sheetData)
        {
            _rules = ParseRulesFromSheet(sheetData);
            System.Diagnostics.Debug.WriteLine($"Loaded {_rules.Count} classification rules");
        }

        /// <summary>
        /// Parse sheet data into ClassificationRule objects
        /// </summary>
        private List<ClassificationRule> ParseRulesFromSheet(IList<IList<object>> sheetData)
        {
            var rules = new List<ClassificationRule>();

            if (sheetData == null || sheetData.Count == 0)
                return rules;

            // Skip header row (row 0), start from row 1
            for (int i = 1; i < sheetData.Count; i++)
            {
                var row = sheetData[i];
                if (row.Count < 6)
                    continue; // Skip incomplete rows

                var rule = new ClassificationRule
                {
                    Scope = GetCellValue(row, 0),
                    ProjectID = GetCellValue(row, 1),
                    TargetValue = GetCellValue(row, 2),
                    MatchText = GetCellValue(row, 3),
                    MatchSender = GetCellValue(row, 4),
                    Priority = int.TryParse(GetCellValue(row, 5), out int priority) ? priority : 999
                };

                rules.Add(rule);
            }

            return rules;
        }

        private string GetCellValue(IList<object> row, int index)
        {
            if (index < row.Count && row[index] != null)
                return row[index].ToString().Trim();
            return string.Empty;
        }

        /// <summary>
        /// Classify an email using Universal Scoring System with back-propagation
        /// </summary>
        public ClassificationResult Classify(string subject, string body, string sender, string toRecipients = "")
        {
            // Normalize inputs
            string searchText = $"{subject} {body}".ToLowerInvariant();
            string normalizedSender = sender?.ToLowerInvariant() ?? string.Empty;

            // Scoring dictionaries
            var projectScores = new Dictionary<string, int>();
            var packageScores = new Dictionary<string, int>();

            // SCORING PHASE: Evaluate ALL rules
            foreach (var rule in _rules)
            {
                int score = 0;

                // Calculate base score from matches
                if (!string.IsNullOrWhiteSpace(rule.MatchText) && MatchesText(searchText, rule.MatchText))
                {
                    score += 100; // Exact text match
                }

                if (!string.IsNullOrWhiteSpace(rule.MatchSender))
                {
                    // Combine FROM sender and TO recipients into a single list for matching
                    var senderList = new List<string> { normalizedSender };
                    var toRecipientsList = ParseRecipients(toRecipients);
                    senderList.AddRange(toRecipientsList);

                    if (MatchesSender(senderList, rule.MatchSender))
                    {
                        score += 50; // Sender or recipient match
                    }
                }

                // Add priority bonus (higher priority = lower number = more points)
                if (score > 0)
                {
                    score += (10 - rule.Priority);
                }

                // Skip if no match
                if (score == 0)
                    continue;

                // Apply score based on rule type
                if (rule.Scope.Equals("PROJECT", StringComparison.OrdinalIgnoreCase))
                {
                    // Add to project scores
                    if (!projectScores.ContainsKey(rule.TargetValue))
                        projectScores[rule.TargetValue] = 0;
                    projectScores[rule.TargetValue] += score;

                    System.Diagnostics.Debug.WriteLine($"Project '{rule.TargetValue}' scored {score} points");
                }
                else if (rule.Scope.Equals("PACKAGE", StringComparison.OrdinalIgnoreCase))
                {
                    // Add to package scores
                    if (!packageScores.ContainsKey(rule.TargetValue))
                        packageScores[rule.TargetValue] = 0;
                    packageScores[rule.TargetValue] += score;

                    // BACK-PROPAGATION: Add score to parent project
                    if (!string.IsNullOrWhiteSpace(rule.ProjectID))
                    {
                        if (!projectScores.ContainsKey(rule.ProjectID))
                            projectScores[rule.ProjectID] = 0;
                        projectScores[rule.ProjectID] += score;

                        System.Diagnostics.Debug.WriteLine($"Package '{rule.TargetValue}' scored {score} points (back-propagated to Project '{rule.ProjectID}')");
                    }
                }
            }

            // SELECTION PHASE: Determine winners and detect ambiguity
            return SelectWinners(projectScores, packageScores);
        }

        private ClassificationResult SelectWinners(Dictionary<string, int> projectScores, Dictionary<string, int> packageScores)
        {
            var result = new ClassificationResult
            {
                SuggestedProjectID = "Random",
                SuggestedPackageID = string.Empty,
                IsAmbiguous = false
            };

            // Sort projects by score descending
            var sortedProjects = projectScores.OrderByDescending(p => p.Value).ToList();
            var sortedPackages = packageScores.OrderByDescending(p => p.Value).ToList();

            // Check Project ambiguity
            if (sortedProjects.Count == 0 || sortedProjects[0].Value == 0)
            {
                result.SuggestedProjectID = "Random";
            }
            else if (sortedProjects.Count > 1 && sortedProjects[0].Value == sortedProjects[1].Value)
            {
                // Ambiguous - multiple projects tied for first
                result.IsAmbiguous = true;
                result.AmbiguityReason = "Multiple projects have matching scores";

                // Add all tied candidates
                int topScore = sortedProjects[0].Value;
                foreach (var project in sortedProjects.Where(p => p.Value == topScore))
                {
                    result.Candidates.Add(new ClassificationCandidate
                    {
                        Name = project.Key,
                        Score = project.Value,
                        Type = "PROJECT"
                    });
                }
            }
            else
            {
                // Clear winner
                result.SuggestedProjectID = sortedProjects[0].Key;
            }

            // Check Package ambiguity
            if (sortedPackages.Count == 0 || sortedPackages[0].Value == 0)
            {
                result.SuggestedPackageID = string.Empty;
            }
            else if (sortedPackages.Count > 1 && sortedPackages[0].Value == sortedPackages[1].Value)
            {
                // Ambiguous - multiple packages tied for first
                result.IsAmbiguous = true;
                if (string.IsNullOrEmpty(result.AmbiguityReason))
                    result.AmbiguityReason = "Multiple packages have matching scores";
                else
                    result.AmbiguityReason += " and multiple packages have matching scores";

                // Add all tied candidates
                int topScore = sortedPackages[0].Value;
                foreach (var package in sortedPackages.Where(p => p.Value == topScore))
                {
                    result.Candidates.Add(new ClassificationCandidate
                    {
                        Name = package.Key,
                        Score = package.Value,
                        Type = "PACKAGE"
                    });
                }
            }
            else
            {
                // Clear winner
                result.SuggestedPackageID = sortedPackages[0].Key;
            }

            System.Diagnostics.Debug.WriteLine($"Classification Result: Project={result.SuggestedProjectID}, Package={result.SuggestedPackageID}, IsAmbiguous={result.IsAmbiguous}");
            if (result.IsAmbiguous)
            {
                System.Diagnostics.Debug.WriteLine($"Ambiguity Reason: {result.AmbiguityReason}, Candidates: {result.Candidates.Count}");
            }

            return result;
        }

        /// <summary>
        /// Check if text contains any of the keywords using word boundary regex
        /// </summary>
        private bool MatchesText(string searchText, string matchText)
        {
            if (string.IsNullOrWhiteSpace(matchText))
                return false;

            // Split by comma and check each keyword
            var keywords = matchText.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var keyword in keywords)
            {
                string trimmed = keyword.Trim();
                if (string.IsNullOrEmpty(trimmed))
                    continue;

                // Use word boundary regex for whole word matching
                string pattern = $@"\b{Regex.Escape(trimmed.ToLowerInvariant())}\b";
                if (Regex.IsMatch(searchText, pattern, RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Check if any sender/recipient email matches any of the sender patterns
        /// </summary>
        private bool MatchesSender(List<string> senderList, string matchSender)
        {
            if (string.IsNullOrWhiteSpace(matchSender))
                return false;

            if (senderList == null || senderList.Count == 0)
                return false;

            // Split by comma and check each sender pattern
            var patterns = matchSender.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var pattern in patterns)
            {
                string trimmed = pattern.Trim().ToLowerInvariant();
                if (string.IsNullOrEmpty(trimmed))
                    continue;

                // Handle wildcard domains: *@brembana.com → @brembana.com
                if (trimmed.StartsWith("*"))
                    trimmed = trimmed.Substring(1);

                // Check if any sender matches this pattern
                foreach (var sender in senderList)
                {
                    if (sender.Contains(trimmed))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Parse semicolon-separated recipient string into individual email addresses
        /// </summary>
        private List<string> ParseRecipients(string recipientString)
        {
            var recipients = new List<string>();

            if (string.IsNullOrWhiteSpace(recipientString))
                return recipients;

            // Split by semicolon (Outlook's standard separator)
            var parts = recipientString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in parts)
            {
                string trimmed = part.Trim().ToLowerInvariant();
                if (!string.IsNullOrEmpty(trimmed))
                {
                    recipients.Add(trimmed);
                }
            }

            return recipients;
        }

        /// <summary>
        /// Get current rule count (for diagnostics)
        /// </summary>
        public int GetRuleCount()
        {
            return _rules.Count;
        }
    }
}
