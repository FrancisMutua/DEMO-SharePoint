namespace DEMO_SharePoint.Services.Models
{
    /// <summary>
    /// Scan job parameters
    /// </summary>
    public class ScanParameters
    {
        public int DPI { get; set; } = 300;
        public string ColorMode { get; set; } = "Grayscale";
        public string PaperSize { get; set; } = "A4";
        public bool Duplex { get; set; } = false;
        public bool UseADF { get; set; } = true;
        public int MaxPages { get; set; } = 0;
        public string OutputFormat { get; set; } = "PDF";
        public string CompressionLevel { get; set; } = "Normal";
    }
}