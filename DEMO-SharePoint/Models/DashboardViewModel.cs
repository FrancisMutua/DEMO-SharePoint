using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DEMO_SharePoint.Models
{
    public class DashboardViewModel
    {
        public int TotalDocuments { get; set; }
        public int UploadsToday { get; set; }
        public int ActiveUsers { get; set; }
        public List<RecentActivity> RecentActivities { get; set; }
    }

    public class RecentActivity
    {
        public string User { get; set; }
        public string Activity { get; set; }
        public string Document { get; set; }
        public DateTime Date { get; set; }
    }

}