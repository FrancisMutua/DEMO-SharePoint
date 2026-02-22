using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DEMO_SharePoint.Services.Models;

namespace DEMO_SharePoint.Services.Interfaces
{
    /// <summary>
    /// Interface for automatic metadata extraction from scanned documents
    /// Identifies and classifies document types, extracts key fields
    /// </summary>
    public interface IMetadataExtractorService
    {
        /// <summary>
        /// Analyzes document and extracts metadata fields
        /// Uses OCR text and barcode data
        /// </summary>
        Task<ExtractedMetadata> ExtractMetadataAsync(
            byte[] documentBytes,
            string ocrText,
            List<BarcodeDetectionResult> barcodes);

        /// <summary>
        /// Classifies document type based on content
        /// Returns confidence scores for top matches
        /// </summary>
        Task<DocumentClassification> ClassifyDocumentAsync(
            string ocrText,
            List<BarcodeDetectionResult> barcodes);

        /// <summary>
        /// Extracts key-value pairs based on templates
        /// </summary>
        Task<Dictionary<string, string>> ExtractKeyValuePairsAsync(
            string ocrText,
            string documentType);

        /// <summary>
        /// Suggests workflow and approval path based on document type
        /// </summary>
        Task<WorkflowSuggestion> SuggestWorkflowAsync(
            string documentType,
            Dictionary<string, string> extractedFields);
    }
}