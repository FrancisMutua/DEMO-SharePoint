using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using DEMO_SharePoint.Models;

namespace DEMO_SharePoint.Models
{
    // Models/WorkflowModel.cs
    public class WorkflowModel
    {
        public int Id { get; set; }

        [Required]
        public string Name { get; set; }

        [Required]
        public string LibraryUrl { get; set; }

        public bool IsActive { get; set; } = true;

        public int Levels { get; set; }

        public List<WorkflowStageModel> Stages { get; set; }
    }
}