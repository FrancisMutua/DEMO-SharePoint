using System.Collections.Generic;

namespace DEMO_SharePoint.Models
{
    /// <summary>
    /// Posted from the "Add Workflow" modal to WorkflowController.CreateWorkflow.
    /// </summary>
    public class CreateWorkflowRequest
    {
        public string Name { get; set; }
        public string LibraryUrl { get; set; }
        public int Levels { get; set; }
        public List<WorkflowStageModel> Stages { get; set; } = new List<WorkflowStageModel>();

        /// <summary>Events that auto-trigger this workflow ("Manual", "Upload", "Update").</summary>
        public List<string> TriggerEvents { get; set; } = new List<string> { "Manual" };

        /// <summary>"ReturnToSubmitter" | "ReturnToPreviousLevel" | "RestartWorkflow"</summary>
        public string RejectionBehavior { get; set; } = "ReturnToSubmitter";

        public bool NotifyOnSubmit { get; set; } = true;
        public bool NotifyOnApprove { get; set; } = true;
        public bool NotifyOnReject { get; set; } = true;
        public bool NotifyOnEscalate { get; set; } = true;
        public bool NotifyOnDelegate { get; set; } = true;
        public bool NotifyOnComplete { get; set; } = true;
    }

    // Action request DTOs (JSON-bound by MVC 5 JsonValueProviderFactory)

    public class ApproveRequest
    {
        public string InstanceId { get; set; }
        public string ItemUrl { get; set; }
        public string Comments { get; set; }
    }

    public class RejectRequest
    {
        public string InstanceId { get; set; }
        public string Comments { get; set; }
    }

    public class DelegateRequest
    {
        public string InstanceId { get; set; }
        public string ToUser { get; set; }
        public string Reason { get; set; }
    }

    public class RecallRequest
    {
        public string WorkflowRunId { get; set; }
    }
}
