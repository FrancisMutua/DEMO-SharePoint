// Models/SPLibrary.cs
using System;

namespace DEMO_SharePoint.Models
{
    public class SPLibrary
    {
       
        public int Id { get; set; }

        public string Title { get; set; }

        public string Url { get; set; }

        public string Description { get; set; }
        public bool IsActive { get; set; } = true;

        public DateTime CreatedOn { get; set; } = DateTime.Now;

        public string Owner { get; set; }
    }
}
