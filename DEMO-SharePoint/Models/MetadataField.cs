using System;
using System.Collections.Generic;

namespace DEMO_SharePoint.Models
{
    /// <summary>
    /// Dynamic metadata field definition for extensible metadata capture
    /// </summary>
    public class MetadataField
    {
        public int Id { get; set; }
        public string FieldDisplayName { get; set; }
        public string FieldName { get; set; }
        public string FieldType { get; set; }
        public bool IsRequired { get; set; }
        public List<string> Choices { get; set; }
        public string TargetLibrary { get; set; }
        public int SequenceOrder { get; set; }
    }
}