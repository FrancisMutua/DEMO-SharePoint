namespace DEMO_SharePoint.Models
{
    /// <summary>
    /// Represents one approval level within a workflow configuration.
    /// Serialised to JSON and stored in WorkflowConfigs.StagesJson.
    /// </summary>
    public class WorkflowStageModel
    {
        public int Level { get; set; }

        /// <summary>
        /// Usernames of all approvers assigned to this level.
        /// </summary>
        public System.Collections.Generic.List<string> Approvers { get; set; }
            = new System.Collections.Generic.List<string>();

        /// <summary>
        /// "Any"  – workflow advances as soon as ONE approver approves.
        ///          All other pending rows at this level become Superseded.
        /// "All"  – every approver at this level must approve before advancing.
        /// </summary>
        public string ApprovalMode { get; set; } = "Any";

        /// <summary>
        /// Number of calendar days from submission/activation before this level
        /// is considered overdue and the escalation contact is notified.
        /// 0 = no escalation.
        /// </summary>
        public int DueInDays { get; set; } = 3;

        /// <summary>
        /// E-mail address (or username) that receives escalation alerts when
        /// DueInDays is exceeded without action.
        /// </summary>
        public string EscalateToEmail { get; set; }

        /// <summary>
        /// Whether approvers at this level may delegate to another user.
        /// </summary>
        public bool AllowDelegation { get; set; } = true;
    }
}
