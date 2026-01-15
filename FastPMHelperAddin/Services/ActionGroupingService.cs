using System;
using System.Collections.Generic;
using System.Linq;
using FastPMHelperAddin.Models;

namespace FastPMHelperAddin.Services
{
    public class ActionGroupingService
    {
        /// <summary>
        /// Groups actions based on email context with priority ordering
        /// </summary>
        public ActionGroupingResult GroupActions(
            List<ActionItem> allActions,
            string internetMessageId,
            string inReplyToId,
            string conversationId,
            string packageContext,
            string projectContext)
        {
            var result = new ActionGroupingResult
            {
                LinkedActions = new List<ActionItem>(),
                PackageActions = new List<ActionItem>(),
                ProjectActions = new List<ActionItem>(),
                OtherActions = new List<ActionItem>()
            };

            if (allActions == null || allActions.Count == 0)
                return result;

            // Track which actions have been categorized (no duplicates)
            var categorized = new HashSet<int>();

            // PASS 1: Linked Actions with weighted scoring
            // Calculate match weight for each action and sort by priority
            var linkedWithWeights = new List<(ActionItem action, int weight)>();

            foreach (var action in allActions)
            {
                int weight = CalculateMatchWeight(action, internetMessageId, inReplyToId, conversationId);
                if (weight > 0)
                {
                    linkedWithWeights.Add((action, weight));
                    categorized.Add(action.Id);
                }
            }

            // Sort by weight descending (highest priority first)
            result.LinkedActions = linkedWithWeights
                .OrderByDescending(x => x.weight)
                .Select(x => x.action)
                .ToList();

            // Determine context from linked actions if available
            if (result.LinkedActions.Count > 0 && string.IsNullOrEmpty(packageContext))
            {
                var firstLinked = result.LinkedActions[0];
                result.DetectedPackage = firstLinked.Package;
                result.DetectedProject = firstLinked.Project;
                packageContext = result.DetectedPackage;
                projectContext = result.DetectedProject;
            }
            else
            {
                result.DetectedPackage = packageContext;
                result.DetectedProject = projectContext;
            }

            // PASS 2: Package Actions
            if (!string.IsNullOrEmpty(packageContext))
            {
                foreach (var action in allActions)
                {
                    if (categorized.Contains(action.Id))
                        continue;

                    if (string.Equals(action.Package, packageContext, StringComparison.OrdinalIgnoreCase))
                    {
                        result.PackageActions.Add(action);
                        categorized.Add(action.Id);
                    }
                }
            }

            // PASS 3: Project Actions (for expanded view)
            if (!string.IsNullOrEmpty(projectContext))
            {
                foreach (var action in allActions)
                {
                    if (categorized.Contains(action.Id))
                        continue;

                    if (string.Equals(action.Project, projectContext, StringComparison.OrdinalIgnoreCase))
                    {
                        result.ProjectActions.Add(action);
                        categorized.Add(action.Id);
                    }
                }
            }

            // PASS 4: Other Actions (remaining)
            foreach (var action in allActions)
            {
                if (!categorized.Contains(action.Id))
                {
                    result.OtherActions.Add(action);
                }
            }

            return result;
        }

        /// <summary>
        /// Calculates match weight for an action based on email properties
        /// Higher weight = higher priority match
        /// </summary>
        private int CalculateMatchWeight(ActionItem action, string internetMessageId, string inReplyToId, string conversationId)
        {
            var activeIds = action.ParseActiveMessageIds();
            var linkedThreads = action.ParseLinkedThreadIds();

            // Normalize inputs
            string normalizedMessageId = NormalizeMessageId(internetMessageId);
            string normalizedReplyToId = NormalizeMessageId(inReplyToId);

            // Priority 1: In-Reply-To matches an ActiveMessageID (direct reply)
            // Weight: 1000 + position from end (last email = highest weight)
            if (!string.IsNullOrEmpty(normalizedReplyToId))
            {
                for (int i = activeIds.Count - 1; i >= 0; i--)
                {
                    var reference = activeIds[i];
                    var parts = reference.Split('|');
                    if (parts.Length >= 3)
                    {
                        string refMessageId = NormalizeMessageId(parts[2]);
                        if (refMessageId == normalizedReplyToId)
                        {
                            // Last ActiveMessageID gets highest weight (1000 + 0)
                            // Earlier messages get lower weight (1000 - position)
                            int positionFromEnd = activeIds.Count - 1 - i;
                            return 1000 + (10 * positionFromEnd);
                        }
                    }
                }
            }

            // Priority 2: This email's InternetMessageId is in ActiveMessageIDs (already tracked)
            if (!string.IsNullOrEmpty(normalizedMessageId))
            {
                for (int i = 0; i < activeIds.Count; i++)
                {
                    var reference = activeIds[i];
                    var parts = reference.Split('|');
                    if (parts.Length >= 3)
                    {
                        string refMessageId = NormalizeMessageId(parts[2]);
                        if (refMessageId == normalizedMessageId)
                        {
                            return 500;
                        }
                    }
                }
            }

            // Priority 3: ConversationID match (lowest priority)
            if (!string.IsNullOrEmpty(conversationId) && linkedThreads.Contains(conversationId))
            {
                return 100;
            }

            return 0; // No match
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
    }

    public class ActionGroupingResult
    {
        public List<ActionItem> LinkedActions { get; set; }
        public List<ActionItem> PackageActions { get; set; }
        public List<ActionItem> ProjectActions { get; set; }
        public List<ActionItem> OtherActions { get; set; }
        public string DetectedPackage { get; set; }
        public string DetectedProject { get; set; }
    }
}
