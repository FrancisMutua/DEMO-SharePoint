using System.Collections.Generic;

namespace DEMO_SharePoint.Services.Models
{
    /// <summary>
    /// Document classification result
    /// </summary>
    public class DocumentClassification
    {
        public string PrimaryType { get; set; }
        public decimal PrimaryConfidence { get; set; }
        public List<ClassificationOption> AlternativeTypes { get; set; } = new List<ClassificationOption>();
        public List<string> DetectedKeywords { get; set; } = new List<string>();
    }

    /// <summary>
    /// Alternative classification option
    /// </summary>
    public class ClassificationOption
    {
        public string DocumentType { get; set; }
        public decimal Confidence { get; set; }
    }
}