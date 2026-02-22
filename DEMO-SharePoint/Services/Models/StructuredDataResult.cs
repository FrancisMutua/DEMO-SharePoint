using System.Collections.Generic;

namespace DEMO_SharePoint.Services.Models
{
    /// <summary>
    /// Structured data extraction (forms, tables)
    /// </summary>
    public class StructuredDataResult
    {
        public Dictionary<string, string> FormFields { get; set; } = new Dictionary<string, string>();
        public List<TableData> Tables { get; set; } = new List<TableData>();
        public List<string> ExtractedEmails { get; set; } = new List<string>();
        public List<string> ExtractedPhoneNumbers { get; set; } = new List<string>();
    }

    /// <summary>
    /// Table data from document
    /// </summary>
    public class TableData
    {
        public int PageNumber { get; set; }
        public List<List<string>> Rows { get; set; } = new List<List<string>>();
    }
}