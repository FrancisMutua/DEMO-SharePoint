using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DEMO_SharePoint.Services.Interfaces;
using DEMO_SharePoint.Services.Models;

namespace DEMO_SharePoint.Services.Implementations
{
    /// <summary>
    /// Implementation for automatic metadata extraction and document classification
    /// Analyzes OCR text and barcodes to classify documents and suggest workflows
    /// </summary>
    public class MetadataExtractorService : IMetadataExtractorService
    {
        private readonly string _extractionEngine;

        public MetadataExtractorService()
        {
            _extractionEngine = ConfigurationManager.AppSettings["MetadataExtractionEngine"] ?? "RuleBasedClassifier";
        }

        public async Task<ExtractedMetadata> ExtractMetadataAsync(
            byte[] documentBytes,
            string ocrText,
            List<BarcodeDetectionResult> barcodes)
        {
            if (documentBytes == null || documentBytes.Length == 0)
                throw new ArgumentException("Document bytes cannot be empty", nameof(documentBytes));

            if (string.IsNullOrEmpty(ocrText))
                ocrText = "No OCR text extracted";

            try
            {
                var classification = await ClassifyDocumentAsync(ocrText, barcodes);
                if (classification == null)
                    throw new Exception("Document classification failed");

                var keyValues = await ExtractKeyValuePairsAsync(ocrText, classification.PrimaryType);
                if (keyValues == null)
                    keyValues = new Dictionary<string, string>();

                var metadata = new ExtractedMetadata
                {
                    ReferenceNumber = ExtractReferenceNumber(ocrText, barcodes),
                    DocumentType = classification.PrimaryType ?? "Document",
                    ClassificationCode = DetermineClassificationCode(classification.PrimaryType),
                    Department = ExtractDepartment(ocrText),
                    RecordDate = ExtractDate(ocrText),
                    Description = ExtractDescription(ocrText),
                    Author = ExtractAuthor(ocrText),
                    CustomFields = keyValues,
                    ExtractionConfidence = classification.PrimaryConfidence,
                    ExtractionWarnings = ValidateExtractedData(classification)
                };

                return await Task.FromResult(metadata);
            }
            catch (Exception ex)
            {
                throw new Exception($"Metadata extraction failed: {ex.Message}", ex);
            }
        }

        public async Task<DocumentClassification> ClassifyDocumentAsync(
            string ocrText,
            List<BarcodeDetectionResult> barcodes)
        {
            if (string.IsNullOrEmpty(ocrText))
                ocrText = string.Empty;

            try
            {
                var classification = new DocumentClassification
                {
                    PrimaryType = ClassifyDocumentType(ocrText),
                    PrimaryConfidence = 85.0m,
                    AlternativeTypes = new List<ClassificationOption>(),
                    DetectedKeywords = ExtractKeywords(ocrText)
                };

                return await Task.FromResult(classification);
            }
            catch (Exception ex)
            {
                throw new Exception($"Document classification failed: {ex.Message}", ex);
            }
        }

        public async Task<Dictionary<string, string>> ExtractKeyValuePairsAsync(
            string ocrText,
            string documentType)
        {
            if (string.IsNullOrEmpty(ocrText))
                return await Task.FromResult(new Dictionary<string, string>());

            try
            {
                var keyValues = new Dictionary<string, string>();

                if (string.IsNullOrEmpty(documentType))
                    documentType = "Document";

                switch (documentType.ToLower())
                {
                    case "invoice":
                        keyValues = ExtractInvoiceData(ocrText);
                        break;
                    case "purchaseorder":
                        keyValues = ExtractPurchaseOrderData(ocrText);
                        break;
                    case "contract":
                        keyValues = ExtractContractData(ocrText);
                        break;
                    case "report":
                        keyValues = ExtractReportData(ocrText);
                        break;
                    default:
                        keyValues = ExtractGenericData(ocrText);
                        break;
                }

                return await Task.FromResult(keyValues ?? new Dictionary<string, string>());
            }
            catch (Exception ex)
            {
                throw new Exception($"Key-value pair extraction failed: {ex.Message}", ex);
            }
        }

        public async Task<WorkflowSuggestion> SuggestWorkflowAsync(
            string documentType,
            Dictionary<string, string> extractedFields)
        {
            try
            {
                if (string.IsNullOrEmpty(documentType))
                    documentType = "Document";

                if (extractedFields == null)
                    extractedFields = new Dictionary<string, string>();

                var suggestion = new WorkflowSuggestion
                {
                    SuggestedWorkflow = GetWorkflowForDocumentType(documentType),
                    SuggestedLibrary = GetLibraryForDocumentType(documentType),
                    SuggestedDepartment = GetDepartmentForDocumentType(documentType),
                    SuggestedApprovers = GetApproversForDocumentType(documentType),
                    RequiresManualReview = NeedsManualReview(extractedFields),
                    ReviewReason = DetermineReviewReason(extractedFields)
                };

                return await Task.FromResult(suggestion);
            }
            catch (Exception ex)
            {
                throw new Exception($"Workflow suggestion failed: {ex.Message}", ex);
            }
        }

        #region Helper Methods

        private string ExtractReferenceNumber(string ocrText, List<BarcodeDetectionResult> barcodes)
        {
            try
            {
                if (barcodes != null && barcodes.Count > 0 && !string.IsNullOrEmpty(barcodes[0].BarcodeValue))
                    return barcodes[0].BarcodeValue;

                if (!string.IsNullOrEmpty(ocrText))
                {
                    var refPattern = @"\b(?:REF|ID|INV|PO)-?(\d{6,12})\b";
                    var match = Regex.Match(ocrText, refPattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                        return match.Value;
                }

                return $"DOC-{Guid.NewGuid().ToString().Substring(0, 8).ToUpper()}";
            }
            catch
            {
                return $"DOC-{Guid.NewGuid().ToString().Substring(0, 8).ToUpper()}";
            }
        }

        private string DetermineClassificationCode(string documentType)
        {
            if (string.IsNullOrEmpty(documentType))
                return "GEN-000";

            switch (documentType.ToLower())
            {
                case "invoice":
                    return "FIN-001";
                case "purchaseorder":
                    return "FIN-002";
                case "contract":
                    return "LEG-001";
                case "report":
                    return "OPS-001";
                default:
                    return "GEN-000";
            }
        }

        private string ExtractDepartment(string ocrText)
        {
            try
            {
                if (string.IsNullOrEmpty(ocrText))
                    return "General";

                var departments = new[] { "Finance", "Operations", "HR", "Sales", "IT" };
                foreach (var dept in departments)
                {
                    if (ocrText.IndexOf(dept, StringComparison.OrdinalIgnoreCase) >= 0)
                        return dept;
                }
                return "General";
            }
            catch
            {
                return "General";
            }
        }

        private DateTime ExtractDate(string ocrText)
        {
            try
            {
                if (string.IsNullOrEmpty(ocrText))
                    return DateTime.Today;

                var datePattern = @"\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})\b";
                var match = Regex.Match(ocrText, datePattern);

                if (match.Success && DateTime.TryParse(match.Value, out var date))
                    return date;

                return DateTime.Today;
            }
            catch
            {
                return DateTime.Today;
            }
        }

        private string ExtractDescription(string ocrText)
        {
            try
            {
                if (string.IsNullOrEmpty(ocrText))
                    return string.Empty;

                var lines = ocrText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                return lines.FirstOrDefault(l => !string.IsNullOrWhiteSpace(l)) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string ExtractAuthor(string ocrText)
        {
            try
            {
                if (string.IsNullOrEmpty(ocrText))
                    return string.Empty;

                var namePattern = @"\b([A-Z][a-z]+ [A-Z][a-z]+)\b";
                var match = Regex.Match(ocrText, namePattern);
                return match.Success ? match.Value : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string ClassifyDocumentType(string ocrText)
        {
            try
            {
                if (string.IsNullOrEmpty(ocrText))
                    return "Document";

                var lowerText = ocrText.ToLower();

                if (lowerText.Contains("invoice") || lowerText.Contains("inv-"))
                    return "Invoice";
                if (lowerText.Contains("purchase order") || lowerText.Contains("po-"))
                    return "PurchaseOrder";
                if (lowerText.Contains("contract") || lowerText.Contains("agreement"))
                    return "Contract";
                if (lowerText.Contains("report"))
                    return "Report";

                return "Document";
            }
            catch
            {
                return "Document";
            }
        }

        private List<string> ExtractKeywords(string ocrText)
        {
            try
            {
                if (string.IsNullOrEmpty(ocrText))
                    return new List<string>();

                var words = ocrText.Split(new[] { ' ', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                return words.Where(w => w.Length > 4).Distinct().Take(10).ToList();
            }
            catch
            {
                return new List<string>();
            }
        }

        private Dictionary<string, string> ExtractInvoiceData(string ocrText)
        {
            var data = new Dictionary<string, string>
            {
                { "Amount", ExtractAmount(ocrText) },
                { "Vendor", ExtractVendor(ocrText) },
                { "DueDate", ExtractDate(ocrText).ToString("yyyy-MM-dd") }
            };
            return data;
        }

        private Dictionary<string, string> ExtractPurchaseOrderData(string ocrText)
        {
            var data = new Dictionary<string, string>
            {
                { "OrderValue", ExtractAmount(ocrText) },
                { "Supplier", ExtractVendor(ocrText) },
                { "DeliveryDate", ExtractDate(ocrText).ToString("yyyy-MM-dd") }
            };
            return data;
        }

        private Dictionary<string, string> ExtractContractData(string ocrText)
        {
            var data = new Dictionary<string, string>
            {
                { "Parties", "To be determined" },
                { "EffectiveDate", ExtractDate(ocrText).ToString("yyyy-MM-dd") },
                { "TermMonths", "12" }
            };
            return data;
        }

        private Dictionary<string, string> ExtractReportData(string ocrText)
        {
            var data = new Dictionary<string, string>
            {
                { "ReportType", "Standard" },
                { "Period", DateTime.Now.Year.ToString() },
                { "Status", "Pending Review" }
            };
            return data;
        }

        private Dictionary<string, string> ExtractGenericData(string ocrText)
        {
            var data = new Dictionary<string, string>
            {
                { "Summary", ExtractDescription(ocrText) }
            };
            return data;
        }

        private string ExtractAmount(string ocrText)
        {
            try
            {
                if (string.IsNullOrEmpty(ocrText))
                    return "0.00";

                var amountPattern = @"(?:[\$£€])?\s*(\d+[.,]\d{2})";
                var match = Regex.Match(ocrText, amountPattern);
                return match.Success ? match.Groups[1].Value : "0.00";
            }
            catch
            {
                return "0.00";
            }
        }

        private string ExtractVendor(string ocrText)
        {
            try
            {
                if (string.IsNullOrEmpty(ocrText))
                    return string.Empty;

                var lines = ocrText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                return lines.FirstOrDefault(l => !string.IsNullOrWhiteSpace(l)) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetWorkflowForDocumentType(string documentType)
        {
            if (string.IsNullOrEmpty(documentType))
                return "StandardApproval";

            switch (documentType.ToLower())
            {
                case "invoice":
                    return "InvoiceApproval";
                case "purchaseorder":
                    return "POApproval";
                case "contract":
                    return "ContractReview";
                default:
                    return "StandardApproval";
            }
        }

        private string GetLibraryForDocumentType(string documentType)
        {
            if (string.IsNullOrEmpty(documentType))
                return "Documents";

            switch (documentType.ToLower())
            {
                case "invoice":
                    return "Invoices";
                case "purchaseorder":
                    return "PurchaseOrders";
                case "contract":
                    return "Contracts";
                default:
                    return "Documents";
            }
        }

        private string GetDepartmentForDocumentType(string documentType)
        {
            if (string.IsNullOrEmpty(documentType))
                return "General";

            switch (documentType.ToLower())
            {
                case "invoice":
                    return "Finance";
                case "purchaseorder":
                    return "Operations";
                case "contract":
                    return "Legal";
                default:
                    return "General";
            }
        }

        private List<string> GetApproversForDocumentType(string documentType)
        {
            if (string.IsNullOrEmpty(documentType))
                return new List<string> { "Manager" };

            switch (documentType.ToLower())
            {
                case "invoice":
                    return new List<string> { "Finance Manager", "Director" };
                case "purchaseorder":
                    return new List<string> { "Procurement", "Manager" };
                default:
                    return new List<string> { "Manager" };
            }
        }

        private bool NeedsManualReview(Dictionary<string, string> extractedFields)
        {
            return extractedFields == null || extractedFields.Count == 0;
        }

        private string DetermineReviewReason(Dictionary<string, string> extractedFields)
        {
            try
            {
                if (extractedFields == null || extractedFields.Count == 0)
                    return "No data extracted";
                if (extractedFields.Count < 2)
                    return "Incomplete data extraction";
                return string.Empty;
            }
            catch
            {
                return "Error determining review reason";
            }
        }

        private List<string> ValidateExtractedData(DocumentClassification classification)
        {
            var warnings = new List<string>();

            try
            {
                if (classification == null)
                {
                    warnings.Add("Classification is null");
                    return warnings;
                }

                if (classification.PrimaryConfidence < 70)
                    warnings.Add("Low confidence classification - manual review recommended");

                if (classification.DetectedKeywords == null || classification.DetectedKeywords.Count == 0)
                    warnings.Add("No relevant keywords detected");
            }
            catch (Exception ex)
            {
                warnings.Add($"Validation error: {ex.Message}");
            }

            return warnings;
        }

        #endregion
    }
}           