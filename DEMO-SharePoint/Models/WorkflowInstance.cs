using System;

namespace DEMO_SharePoint.Models
{
    /// <summary>
    /// One row in the ApprovalInstances SharePoint list.
    /// Each row represents one approver's assignment for one stage
    /// of one workflow run (submission).
    /// </summary>
    public class WorkflowInstance
    {
        // SharePoint item ID
        public string Id { get; set; }

        // Run identity
        /// <summary>
        /// GUID shared by all instances that belong to the same document
        /// submission. Allows grouping all stages of a single run.
        /// </summary>
        public string WorkflowRunId { get; set; }

        // Document reference
        public string ItemUrl { get; set; }
        public string ItemName { get; set; }

        // Workflow reference
        public string WorkflowId { get; set; }
        public string WorkflowName { get; set; }

        // Stage info
        public string CurrentLevel { get; set; }
        public string TotalLevels { get; set; }

        /// <summary>
        /// "Any" or "All" – copied from WorkflowStageModel.ApprovalMode.
        /// </summary>
        public string ApprovalMode { get; set; } = "Any";

        // Status
        /// <summary>
        /// Waiting    – not yet this approver's turn (previous stage in progress).
        /// Pending    – awaiting this approver's action.
        /// Approved   – this approver approved.
        /// Rejected   – this approver rejected.
        /// Delegated  – this approver delegated to someone else.
        /// Superseded – another approver in "Any" mode already resolved this level.
        /// Cancelled  – workflow recalled or rejected upstream.
        /// </summary>
        public string Status { get; set; } = "Pending";

        // People
        public string Approver { get; set; }

        /// <summary>Original approver username when delegation has occurred.</summary>
        public string OriginalApprover { get; set; }

        public string SubmittedBy { get; set; }

        // Dates
        public DateTime SubmittedDate { get; set; } = DateTime.Now;
        public DateTime? ActionDate { get; set; }

        /// <summary>
        /// Deadline for action at this stage. Calculated from WorkflowStageModel.DueInDays.
        /// </summary>
        public DateTime? DueDate { get; set; }

        // Comments / audit
        public string Comments { get; set; }

        /// <summary>True once an escalation notification has been sent for this row.</summary>
        public bool IsEscalated { get; set; }

        // Computed helpers (not persisted)
        public bool IsOverdue => DueDate.HasValue && DueDate.Value < DateTime.Now
                                 && Status == "Pending";

        public bool IsPendingAction => Status == "Pending";
    }
}
