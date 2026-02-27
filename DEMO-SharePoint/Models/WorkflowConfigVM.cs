using System.Collections.Generic;
using System.Web.Mvc;

namespace DEMO_SharePoint.Models
{
    /// <summary>
    /// Represents a workflow configuration record from the WorkflowConfigs SP list.
    /// Also carries UI drop-down data for the management screen.
    /// </summary>
    public class WorkflowConfigVM
    {
        // Identity
        public int Id { get; set; }
        public string WorkflowName { get; set; }

        /// <summary>Server-relative URL of the document library this workflow governs.</summary>
        public string LibraryUrl { get; set; }

        public int Levels { get; set; }
        public bool IsActive { get; set; } = true;

        // Stage definitions
        public List<WorkflowStageModel> Stages { get; set; } = new List<WorkflowStageModel>();

        // Trigger configuration
        /// <summary>
        /// Events that auto-start this workflow.
        /// Values: "Manual" | "Upload" | "Update"
        /// Stored as JSON array in SP list field TriggerEvents.
        /// </summary>
        public List<string> TriggerEvents { get; set; } = new List<string> { "Manual" };

        // Rejection behaviour
        /// <summary>
        /// What happens when an approver rejects a document:
        /// "ReturnToSubmitter"    – cancel all pending rows, notify submitter.
        /// "ReturnToPreviousLevel"– re-activate the previous stage approvers.
        /// "RestartWorkflow"      – cancel all, recreate from Level 1.
        /// </summary>
        public string RejectionBehavior { get; set; } = "ReturnToSubmitter";

        // Notification flags
        public bool NotifyOnSubmit { get; set; } = true;
        public bool NotifyOnApprove { get; set; } = true;
        public bool NotifyOnReject { get; set; } = true;
        public bool NotifyOnEscalate { get; set; } = true;
        public bool NotifyOnDelegate { get; set; } = true;
        public bool NotifyOnComplete { get; set; } = true;

        // UI helpers (not persisted)
        public List<SPLibrary> Libraries { get; set; } = new List<SPLibrary>();
        public List<SelectListItem> ApprovalLevels { get; set; } = new List<SelectListItem>();
        public List<WorkflowConfigVM> ExistingWorkflows { get; set; } = new List<WorkflowConfigVM>();
    }
}
