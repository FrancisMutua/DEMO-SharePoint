using System;

namespace DEMO_SharePoint.Models
{
    /// <summary>
    /// Reusable scan profile for quick scanning with preset parameters
    /// </summary>
    public class ScanProfile
    {
        public int Id { get; set; }
        public string ProfileName { get; set; }
        public int DPI { get; set; }
        public string ColorMode { get; set; }
        public bool Duplex { get; set; }
        public string PaperSize { get; set; }
        public bool UseADF { get; set; }
        public string DefaultLibrary { get; set; }
        public DateTime CreatedDate { get; set; }
        public string CreatedBy { get; set; }
    }
}