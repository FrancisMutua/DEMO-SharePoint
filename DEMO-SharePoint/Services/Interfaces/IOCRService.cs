using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DEMO_SharePoint.Services.Models;

namespace DEMO_SharePoint.Services.Interfaces
{
    /// <summary>
    /// Interface for Optical Character Recognition (OCR) processing
    /// Extracts text content from scanned documents
    /// </summary>
    public interface IOCRService
    {
        /// <summary>
        /// Extracts text from scanned document bytes
        /// Supports PDF and image formats
        /// </summary>
        Task<string> ExtractTextAsync(byte[] documentBytes, string fileFormat = "PDF");

        /// <summary>
        /// Extracts text with confidence scores for AI validation
        /// </summary>
        Task<OCRResult> ExtractTextWithConfidenceAsync(byte[] documentBytes, string fileFormat = "PDF");

        /// <summary>
        /// Detects and extracts barcodes from scanned pages
        /// </summary>
        Task<List<BarcodeDetectionResult>> DetectBarcodesAsync(byte[] documentBytes);

        /// <summary>
        /// Extracts structured data (tables, forms) from documents
        /// </summary>
        Task<StructuredDataResult> ExtractStructuredDataAsync(byte[] documentBytes);

        /// <summary>
        /// Detects document language for multi-language support
        /// </summary>
        Task<string> DetectLanguageAsync(byte[] documentBytes);
    }
}