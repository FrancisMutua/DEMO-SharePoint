namespace DEMO_SharePoint.Services.Models
{
    /// <summary>
    /// Barcode detection result
    /// </summary>
    public class BarcodeDetectionResult
    {
        public string BarcodeValue { get; set; }
        public string BarcodeFormat { get; set; }
        public int PageNumber { get; set; }
        public decimal Confidence { get; set; }
    }
}