using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace KCAU_SharePoint.Models
{
    public class WorkflowApprovalHistory
    {
        public int Id { get; set; }
        public int WorkflowInstanceId { get; set; }
        public int Level { get; set; }
        public string Approver { get; set; }
        public string Action { get; set; } // Approved / Rejected
        public string Comments { get; set; }
        public DateTime ActionDate { get; set; } = DateTime.Now;
    }
}