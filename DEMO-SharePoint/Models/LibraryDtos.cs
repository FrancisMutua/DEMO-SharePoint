using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace KCAU_SharePoint.Models
{
    public class LibraryDto
    {
        public string Name { get; set; }
        public string Url { get; set; }
        public string Icon { get; set; }
        public string Access { get; set; }
        public string Desc { get; set; }
        public List<MetaFieldDto> Fields { get; set; }
    }

    public class MetaFieldDto
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public bool Required { get; set; }
    }
}