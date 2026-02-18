using System;

namespace KCAU_SharePoint.Models
{
    public class WorkflowInstance
    {
        public string Id { get; set; }
        public string ItemUrl { get; set; } // file or folder
        public string ItemName { get; set; }
        public string WorkflowId { get; set; }
        public string CurrentLevel { get; set; }
        public string Status { get; set; } = "Pending"; // Pending, Approved, Rejected
        public string SubmittedBy { get; set; }
        public DateTime SubmittedDate { get; set; } = DateTime.Now;
        public DateTime? CompletedDate { get; set; }
        public string TotalLevels { get; set; }

        // Only one approver per stage
        public string Approver { get; set; }
        public string Comments { get; set; }
    }
}
