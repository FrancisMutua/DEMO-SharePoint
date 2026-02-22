using System;
using System.Collections.Generic;

namespace DEMO_SharePoint.Services.Models
{
    /// <summary>
    /// OCR extraction result with text and metadata
    /// </summary>
    public class OCRResult
    {
        public string ExtractedText { get; set; }
        public decimal AverageConfidence { get; set; }
        public List<OCRLine> Lines { get; set; } = new List<OCRLine>();
        public int PageCount { get; set; }
        public string DetectedLanguage { get; set; }
        public DateTime ProcessedAt { get; set; }
    }

    /// <summary>
    /// Individual OCR line with confidence
    /// </summary>
    public class OCRLine
    {
        public int PageNumber { get; set; }
        public string Text { get; set; }
        public decimal Confidence { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }
}