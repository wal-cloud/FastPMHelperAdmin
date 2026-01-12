using System;
using System.Collections.Generic;

namespace FastPMHelperAddin.Models
{
    public class ClassificationRule
    {
        public string Scope { get; set; }           // "PROJECT" or "PACKAGE"
        public string ProjectID { get; set; }       // Parent project ID
        public string TargetValue { get; set; }     // The ID to assign
        public string MatchText { get; set; }       // Comma-separated keywords
        public string MatchSender { get; set; }     // Comma-separated sender domains
        public int Priority { get; set; }           // 1 = highest priority
    }

    public class ClassificationResult
    {
        public string SuggestedProjectID { get; set; }
        public string SuggestedPackageID { get; set; }
        public bool IsAmbiguous { get; set; }
        public string AmbiguityReason { get; set; }
        public List<ClassificationCandidate> Candidates { get; set; }

        public ClassificationResult()
        {
            Candidates = new List<ClassificationCandidate>();
        }
    }

    public class ClassificationCandidate
    {
        public string Name { get; set; }        // ProjectID or PackageID
        public int Score { get; set; }
        public string Type { get; set; }        // "PROJECT" or "PACKAGE"
    }
}
