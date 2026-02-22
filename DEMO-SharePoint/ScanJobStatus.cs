using System;

namespace DEMO_SharePoint.Services.Models
{
    /// <summary>
    /// Scan job status response
    /// </summary>
    public class ScanJobStatus
    {
        public string JobId { get; set; }
        public string Status { get; set; }
        public int PagesScanned { get; set; }
        public int EstimatedTotalPages { get; set; }
        public decimal ProgressPercentage => EstimatedTotalPages > 0 ? (PagesScanned * 100M / EstimatedTotalPages) : 0;
        public string ErrorMessage { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime LastUpdated { get; set; }
    }
}