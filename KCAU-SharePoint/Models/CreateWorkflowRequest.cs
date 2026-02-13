using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace KCAU_SharePoint.Models
{
    // Models/CreateWorkflowRequest.cs
    public class CreateWorkflowRequest
    {
        public string Name { get; set; }
        public string LibraryUrl { get; set; }
        public int Levels { get; set; }
        public List<WorkflowStageModel> Stages { get; set; }
    }

}