using System.Collections.Generic;

namespace DEMO_SharePoint.Services.Models
{
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