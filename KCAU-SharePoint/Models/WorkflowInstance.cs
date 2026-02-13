using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace KCAU_SharePoint.Models
{
    public class WorkflowInstance
    {
        public int Id { get; set; }
        public string ItemUrl { get; set; } // file or folder
        public string ItemName { get; set; }
        public int WorkflowId { get; set; }
        public int CurrentLevel { get; set; } = 1;
        public string Status { get; set; } = "Pending"; // Pending, Approved, Rejected
        public string SubmittedBy { get; set; }
        public DateTime SubmittedDate { get; set; } = DateTime.Now;
        public DateTime? CompletedDate { get; set; }
        public int TotalLevels { get; set; }
    }
}