using System;
using System.Collections.Generic;
using System.Web.Mvc;
using DEMO_SharePoint.Models;

namespace DEMO_SharePoint.Models
{
    public class WorkflowConfigVM
    {
        public int Id { get; set; } // SharePoint item ID
        public string WorkflowName { get; set; } // Workflow title
        public string LibraryUrl { get; set; } // Selected library URL
        public int Levels { get; set; } // Number of approval levels
        public bool IsActive { get; set; } = true; // Active status

        public List<WorkflowStageModel> Stages { get; set; } = new List<WorkflowStageModel>(); // Approval stages

        // Dropdowns / selections
        public List<SPLibrary> Libraries { get; set; } = new List<SPLibrary>();
        public List<SelectListItem> ApprovalLevels { get; set; } = new List<SelectListItem>();

        // Optional: list of existing workflows for display
        public List<WorkflowConfigVM> ExistingWorkflows { get; set; } = new List<WorkflowConfigVM>();
    }
}
