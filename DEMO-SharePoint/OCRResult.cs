using System;
using System.Collections.Generic;

namespace DEMO_SharePoint.Services.Models
{
    /// <summary>
    /// Extracted metadata from document
    /// </summary>
    public class ExtractedMetadata
    {
        public string ReferenceNumber { get; set; }
        public string DocumentType { get; set; }
        public string ClassificationCode { get; set; }
        public string Department { get; set; }
        public DateTime RecordDate { get; set; }
        public string Description { get; set; }
        public string Author { get; set; }
        public Dictionary<string, string> CustomFields { get; set; } = new Dictionary<string, string>();
        public decimal ExtractionConfidence { get; set; }
        public List<string> ExtractionWarnings { get; set; } = new List<string>();
    }

    /// <summary>
    /// Suggested workflow and routing
    /// </summary>
    public class WorkflowSuggestion
    {
        public string SuggestedWorkflow { get; set; }
        public string SuggestedLibrary { get; set; }
        public string SuggestedDepartment { get; set; }
        public List<string> SuggestedApprovers { get; set; } = new List<string>();
        public bool RequiresManualReview { get; set; }
        public string ReviewReason { get; set; }
    }
}