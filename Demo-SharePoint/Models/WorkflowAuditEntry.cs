using System;

namespace DEMO_SharePoint.Models
{
    /// <summary>
    /// One row in the WorkflowAuditLog SharePoint list.
    /// Provides an immutable, chronological record of every action taken
    /// within a workflow run.
    /// </summary>
    public class WorkflowAuditEntry
    {
        public string Id { get; set; }

        /// <summary>Links all audit entries for a single document submission.</summary>
        public string WorkflowRunId { get; set; }

        public string ItemUrl { get; set; }
        public string ItemName { get; set; }
        public string WorkflowName { get; set; }

        /// <summary>Username who performed the action.</summary>
        public string Actor { get; set; }

        /// <summary>
        /// One of: Submitted | Approved | Rejected | Delegated | Escalated |
        ///         Recalled | Completed | StageAdvanced | Cancelled | Restarted
        /// </summary>
        public string Action { get; set; }

        public int Stage { get; set; }
        public string Comments { get; set; }
        public DateTime Timestamp { get; set; }

        // Display helpers
        public string BadgeClass
        {
            get
            {
                switch (Action)
                {
                    case "Submitted":     return "bg-primary";
                    case "Approved":      return "bg-success";
                    case "StageAdvanced": return "bg-info text-dark";
                    case "Completed":     return "bg-success";
                    case "Rejected":      return "bg-danger";
                    case "Recalled":      return "bg-warning text-dark";
                    case "Delegated":     return "bg-secondary";
                    case "Escalated":     return "bg-warning text-dark";
                    case "Cancelled":     return "bg-secondary";
                    case "Restarted":     return "bg-info text-dark";
                    default:              return "bg-secondary";
                }
            }
        }

        public string Icon
        {
            get
            {
                switch (Action)
                {
                    case "Submitted":     return "bi-send";
                    case "Approved":      return "bi-check-circle-fill";
                    case "StageAdvanced": return "bi-arrow-right-circle";
                    case "Completed":     return "bi-patch-check-fill";
                    case "Rejected":      return "bi-x-circle-fill";
                    case "Recalled":      return "bi-arrow-counterclockwise";
                    case "Delegated":     return "bi-person-fill-add";
                    case "Escalated":     return "bi-exclamation-triangle-fill";
                    case "Cancelled":     return "bi-slash-circle";
                    case "Restarted":     return "bi-arrow-repeat";
                    default:              return "bi-dot";
                }
            }
        }
    }
}
