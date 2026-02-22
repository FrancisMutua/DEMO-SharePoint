using System;
using System.Collections.Generic;

namespace DEMO_SharePoint.Services.Models
{
    /// <summary>
    /// Network scanner device representation
    /// </summary>
    public class NetworkScanner
    {
        public string ScannerId { get; set; }
        public string FriendlyName { get; set; }
        public string IPAddress { get; set; }
        public string Manufacturer { get; set; }
        public string Model { get; set; }
        public string SerialNumber { get; set; }
        public bool IsOnline { get; set; }
        public List<string> SupportedColorModes { get; set; } = new List<string> { "Color", "Grayscale", "BlackWhite" };
        public List<int> SupportedDPIs { get; set; } = new List<int> { 150, 200, 300, 600 };
        public List<string> SupportedPaperSizes { get; set; } = new List<string> { "A4", "A3", "Letter", "Legal" };
        public bool SupportsDuplex { get; set; }
        public bool SupportsADF { get; set; }
        public DateTime LastStatusCheck { get; set; }
    }
}