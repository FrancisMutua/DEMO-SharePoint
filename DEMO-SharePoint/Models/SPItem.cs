using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DEMO_SharePoint.Models
{
    public class SPItem
    {
        public string Name { get; set; }
        public string Url { get; set; }
        public bool IsFolder { get; set; }
        public Dictionary<string, string> Metadata { get; set; } = new Dictionary<string, string>();
        public string LibraryUrl { get; set; }
    }
}