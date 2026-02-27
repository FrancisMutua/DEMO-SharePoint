using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using DEMO_SharePoint.Models;
using DEMO_SharePoint.Services.Implementations;
using DEMO_SharePoint.Services.Interfaces;
using DEMO_SharePoint.Services.Models;
using Microsoft.SharePoint.Client;

namespace DEMO_SharePoint.Controllers
{
    /// <summary>
    /// Complete network document scanning workflow
    /// Handles scanner discovery, OCR, metadata extraction, and SharePoint upload
    /// </summary>
    [SessionAuthorize]
    public class ScanController : Controller
    {
        private readonly INetworkPrinterService _scannerService;
        private readonly IOCRService _ocrService;
        private readonly IMetadataExtractorService _metadataService;
        private readonly Helper _helper;

        /// <summary>
        /// Parameterless constructor for MVC framework instantiation
        /// </summary>
        public ScanController() 
            : this(null, null, null, null)
        {
        }

        /// <summary>
        /// Constructor with dependency injection support
        /// Services use default implementations if not provided
        /// </summary>
        public ScanController(
            INetworkPrinterService scannerService = null,
            IOCRService ocrService = null,
            IMetadataExtractorService metadataService = null,
            Helper helper = null)
        {
            _scannerService = scannerService ?? new NetworkPrinterService();
            _ocrService = ocrService ?? new OCRService();
            _metadataService = metadataService ?? new MetadataExtractorService();
            _helper = helper ?? new Helper();
        }

        /// <summary>
        /// STEP 1 & 2: Scanner selection and scan settings interface
        /// Displays available network scanners and scan parameters
        /// </summary>
        [HttpGet]
        public async Task<ActionResult> Index()
        {
            try
            {
                var scanners = await _scannerService.DiscoverNetworkScannersAsync();
                var profiles = GetScanProfiles();
                var libraries = GetTargetLibraries();

                var vm = new ScanIndexViewModel
                {
                    AvailableScanners = scanners?.Select(s => s.FriendlyName).ToList() ?? new List<string>(),
                    ScanProfiles = profiles ?? new List<ScanProfile>(),
                    LibraryOptions = libraries ?? new List<SelectListItem>(),
                    AgentOnline = scanners != null && scanners.Count > 0,
                    AgentErrorMessage = scanners == null ? "No scanners found" : null
                };

                return View(vm);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Scanner discovery error: {ex.Message}");
                return View(new ScanIndexViewModel
                {
                    AgentOnline = false,
                    AgentErrorMessage = $"Error: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// STEP 2B: Initiate scan job on selected network scanner
        /// Returns job ID for progress tracking
        /// </summary>
        [HttpPost]
        public async Task<JsonResult> InitiateScan(ScanSettingsViewModel settings)
        {
            try
            {
                if (settings == null)
                    return Json(new { success = false, message = "Settings cannot be null" });

                if (string.IsNullOrWhiteSpace(settings.SelectedScanner))
                    return Json(new { success = false, message = "Please select a scanner" });

                var scanParams = new ScanParameters
                {
                    DPI = settings.DPI > 0 ? settings.DPI : 300,
                    ColorMode = !string.IsNullOrEmpty(settings.ColorMode) ? settings.ColorMode : "Grayscale",
                    PaperSize = !string.IsNullOrEmpty(settings.PaperSize) ? settings.PaperSize : "A4",
                    Duplex = settings.Duplex,
                    UseADF = settings.UseADF
                };

                var scanners = await _scannerService.DiscoverNetworkScannersAsync();
                if (scanners == null || scanners.Count == 0)
                    return Json(new { success = false, message = "No scanners available" });

                var scanner = scanners.FirstOrDefault(s => 
                    s.FriendlyName.Equals(settings.SelectedScanner, StringComparison.OrdinalIgnoreCase));

                if (scanner == null)
                    return Json(new { success = false, message = "Selected scanner not found" });

                var jobId = await _scannerService.InitiateScanAsync(scanner.ScannerId, scanParams);

                if (string.IsNullOrEmpty(jobId))
                    return Json(new { success = false, message = "Failed to initiate scan job" });

                Session["ScanJobId"] = jobId;
                Session["ScannerName"] = settings.SelectedScanner;

                return Json(new { success = true, jobId = jobId });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"InitiateScan error: {ex.Message}");
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Monitor scan job progress in real-time
        /// Returns page count and completion percentage
        /// </summary>
        [HttpGet]
        public async Task<JsonResult> GetScanProgress(string jobId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(jobId))
                    return Json(new { success = false, message = "Job ID is required" }, JsonRequestBehavior.AllowGet);

                var status = await _scannerService.GetScanStatusAsync(jobId);
                
                if (status == null)
                    return Json(new { success = false, message = "Status not found" }, JsonRequestBehavior.AllowGet);

                return Json(new
                {
                    success = true,
                    status = status.Status ?? "Unknown",
                    pagesScanned = status.PagesScanned,
                    estimatedTotal = status.EstimatedTotalPages,
                    progress = status.ProgressPercentage,
                    error = status.ErrorMessage
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetScanProgress error: {ex.Message}");
                return Json(new { success = false, message = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// STEP 3: Download completed scan and extract OCR + metadata
        /// Automatically classifies document and suggests workflow
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> ProcessScan(string jobId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(jobId))
                    throw new Exception("Job ID is required");

                var documentBytes = await _scannerService.DownloadScanAsync(jobId);
                
                if (documentBytes == null || documentBytes.Length == 0)
                    throw new Exception("Failed to download scan - no data received");

                // Extract OCR text
                var ocrResult = await _ocrService.ExtractTextWithConfidenceAsync(documentBytes);
                
                if (ocrResult == null)
                    throw new Exception("OCR extraction returned null result");

                var ocrText = ocrResult.ExtractedText ?? string.Empty;

                // Detect barcodes
                var barcodes = await _ocrService.DetectBarcodesAsync(documentBytes);
                if (barcodes == null)
                    barcodes = new List<BarcodeDetectionResult>();

                // Extract metadata
                var extractedMetadata = await _metadataService.ExtractMetadataAsync(
                    documentBytes,
                    ocrText,
                    barcodes);

                if (extractedMetadata == null)
                    throw new Exception("Metadata extraction returned null result");

                // Suggest workflow
                var workflowSuggestion = await _metadataService.SuggestWorkflowAsync(
                    extractedMetadata.DocumentType ?? "Document",
                    extractedMetadata.CustomFields ?? new Dictionary<string, string>());

                if (workflowSuggestion == null)
                    throw new Exception("Workflow suggestion returned null result");

                // Store in session for next steps
                Session["ExtractedMetadata"] = extractedMetadata;
                Session["WorkflowSuggestion"] = workflowSuggestion;
                Session["ScanBytes"] = documentBytes;
                Session["BarcodeData"] = barcodes;
                Session["OCRText"] = ocrText;

                return RedirectToAction("Preview", new { jobId = jobId });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProcessScan error: {ex.Message}");
                return View("Error", new HandleErrorInfo(ex, "Scan", "ProcessScan"));
            }
        }

        /// <summary>
        /// STEP 3: Document preview with OCR and barcode results
        /// Shows extracted text and detected barcodes for review
        /// </summary>
        [HttpGet]
        public ActionResult Preview(string jobId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(jobId))
                    throw new Exception("Job ID is required");

                var extractedMetadata = Session["ExtractedMetadata"] as ExtractedMetadata;
                var ocrText = (Session["OCRText"] ?? "").ToString();
                var barcodes = Session["BarcodeData"] as List<BarcodeDetectionResult>;

                var vm = new PreviewViewModel
                {
                    PageCount = 1,
                    OCRText = !string.IsNullOrEmpty(ocrText) ? ocrText : "No text extracted",
                    BarcodeValue = barcodes?.FirstOrDefault()?.BarcodeValue,
                    BarcodeFormat = barcodes?.FirstOrDefault()?.BarcodeFormat
                };

                return View(vm);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Preview error: {ex.Message}");
                return View("Error", new HandleErrorInfo(ex, "Scan", "Preview"));
            }
        }

        /// <summary>
        /// STEP 4: Metadata entry form
        /// Pre-populated with extracted metadata, user can edit before upload
        /// </summary>
        [HttpGet]
        public ActionResult Metadata()
        {
            try
            {
                var extractedMetadata = Session["ExtractedMetadata"] as ExtractedMetadata;
                var workflowSuggestion = Session["WorkflowSuggestion"] as WorkflowSuggestion;

                var vm = new MetadataViewModel
                {
                    ReferenceNumber = extractedMetadata?.ReferenceNumber ?? $"DOC-{DateTime.Now:yyyyMMddHHmmss}",
                    DocumentType = extractedMetadata?.DocumentType ?? "Document",
                    ClassificationCode = extractedMetadata?.ClassificationCode ?? "GEN-000",
                    Department = extractedMetadata?.Department ?? "General",
                    RecordDate = extractedMetadata?.RecordDate ?? DateTime.Today,
                    Description = extractedMetadata?.Description,
                    Author = extractedMetadata?.Author,
                    TargetLibrary = workflowSuggestion?.SuggestedLibrary ?? "Documents",
                    WorkflowName = workflowSuggestion?.SuggestedWorkflow,
                    OCRText = (Session["OCRText"] ?? "").ToString(),
                    PageCount = 1,
                    DynamicFieldValues = extractedMetadata?.CustomFields ?? new Dictionary<string, string>(),
                    ClassificationOptions = GetClassificationOptions(),
                    DocumentTypeOptions = GetDocumentTypeOptions(),
                    DepartmentOptions = GetDepartmentOptions(),
                    LibraryOptions = GetTargetLibraries(),
                    WorkflowOptions = GetWorkflowOptions()
                };

                return View(vm);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Metadata error: {ex.Message}");
                return View("Error", new HandleErrorInfo(ex, "Scan", "Metadata"));
            }
        }

        /// <summary>
        /// STEP 5: Upload to SharePoint + trigger workflow
        /// Final step - uploads document with metadata and optionally starts approval workflow
        /// </summary>
        [HttpPost]
        [ValidateInput(false)]
        public async Task<ActionResult> UploadDocument(MetadataViewModel metadata)
        {
            try
            {
                if (metadata == null)
                    throw new Exception("Metadata cannot be null");

                if (!ModelState.IsValid)
                    return View("Metadata", metadata);

                var documentBytes = Session["ScanBytes"] as byte[];
                if (documentBytes == null || documentBytes.Length == 0)
                    throw new Exception("No scanned document found in session");

                var fileName = GenerateDocumentFileName(metadata);

                using (var ctx = _helper.GetContext())
                {
                    var web = ctx.Web;
                    var lib = web.GetList($"/sites/edms/{metadata.TargetLibrary}");

                    // Upload file
                    var fileCreationInfo = new FileCreationInformation
                    {
                        Content = documentBytes,
                        Url = fileName,
                        Overwrite = false
                    };

                    var uploadFile = lib.RootFolder.Files.Add(fileCreationInfo);
                    var listItem = uploadFile.ListItemAllFields;

                    // Set metadata
                    listItem["Title"] = fileName;
                    listItem["ReferenceNumber"] = metadata.ReferenceNumber;
                    listItem["DocumentType"] = metadata.DocumentType;
                    listItem["ClassificationCode"] = metadata.ClassificationCode;
                    listItem["Department"] = metadata.Department;
                    listItem["RecordDate"] = metadata.RecordDate;
                    listItem["Description"] = metadata.Description;
                    listItem["Author"] = metadata.Author ?? HttpContext.Session["Username"]?.ToString();
                    
                    if (!string.IsNullOrEmpty(metadata.ReferenceNumber))
                        listItem["ReferenceNumber"] = metadata.ReferenceNumber;
                    
                    if (!string.IsNullOrEmpty(metadata.DocumentType))
                        listItem["DocumentType"] = metadata.DocumentType;
                    
                    if (!string.IsNullOrEmpty(metadata.ClassificationCode))
                        listItem["ClassificationCode"] = metadata.ClassificationCode;
                    
                    if (!string.IsNullOrEmpty(metadata.Department))
                        listItem["Department"] = metadata.Department;
                    
                    if (metadata.RecordDate != DateTime.MinValue)
                        listItem["RecordDate"] = metadata.RecordDate;
                    
                    if (!string.IsNullOrEmpty(metadata.Description))
                        listItem["Description"] = metadata.Description;
                    
                    var author = metadata.Author ?? HttpContext.Session["Username"]?.ToString() ?? "System";
                    if (!string.IsNullOrEmpty(author))
                        listItem["Author"] = author;
                    
                    // Custom fields
                    if (metadata.DynamicFieldValues != null && metadata.DynamicFieldValues.Count > 0)
                    {
                        foreach (var kvp in metadata.DynamicFieldValues)
                        {
                            try 
                            { 
                                if (!string.IsNullOrEmpty(kvp.Key) && !string.IsNullOrEmpty(kvp.Value))
                                    listItem[kvp.Key] = kvp.Value; 
                            }
                            catch (Exception fieldEx)
                            { 
                                System.Diagnostics.Debug.WriteLine($"Field '{kvp.Key}' error: {fieldEx.Message}");
                            }
                        }
                    }

                    listItem.Update();
                    ctx.ExecuteQuery();

                    var itemId = uploadFile.ListItemAllFields.Id;

                    // Trigger workflow if requested
                    if (metadata.SendToWorkflow && !string.IsNullOrEmpty(metadata.WorkflowName))
                    {
                        try
                        {
                            var itemUrl = uploadFile.ServerRelativeUrl;
                            var username = HttpContext.Session["Username"]?.ToString() ?? "Unknown";
                            _helper.ProcessTrigger(itemUrl, fileName, username, "Upload");
                        }
                        catch (Exception workflowEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"Workflow creation error: {workflowEx.Message}");
                        }
                    }

                    var result = new UploadSuccessViewModel
                    {
                        ItemId = itemId,
                        FileName = fileName,
                        Library = metadata.TargetLibrary,
                        FileUrl = uploadFile.ServerRelativeUrl,
                        WorkflowStarted = metadata.SendToWorkflow,
                        Archived = metadata.ArchiveImmediately
                    };

                    // Clean up session
                    Session.Remove("ScanBytes");
                    Session.Remove("ExtractedMetadata");
                    Session.Remove("WorkflowSuggestion");
                    Session.Remove("OCRText");
                    Session.Remove("BarcodeData");
                    Session.Remove("ScanJobId");
                    Session.Remove("ScannerName");

                    return View("UploadSuccess", result);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UploadDocument error: {ex.Message}");
                return View("Error", new HandleErrorInfo(ex, "Scan", "UploadDocument"));
            }
        }

        /// <summary>
        /// Cancel an active scan job
        /// </summary>
        [HttpPost]
        public async Task<JsonResult> CancelScan(string jobId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(jobId))
                    return Json(new { success = false, message = "Job ID is required" });

                var result = await _scannerService.CancelScanAsync(jobId);
                return Json(new { success = result });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        #region Helper Methods

        /// <summary>
        /// Get default scan profiles
        /// </summary>
        private List<ScanProfile> GetScanProfiles()
        {
            return new List<ScanProfile>
            {
                new ScanProfile 
                { 
                    Id = 1, 
                    ProfileName = "Standard (300 DPI)", 
                    DPI = 300, 
                    ColorMode = "Grayscale", 
                    Duplex = false, 
                    PaperSize = "A4",
                    UseADF = true
                },
                new ScanProfile 
                { 
                    Id = 2, 
                    ProfileName = "High Quality (600 DPI)", 
                    DPI = 600, 
                    ColorMode = "Color", 
                    Duplex = true, 
                    PaperSize = "A4",
                    UseADF = true
                },
                new ScanProfile 
                { 
                    Id = 3, 
                    ProfileName = "Fast Scan (150 DPI)", 
                    DPI = 150, 
                    ColorMode = "Grayscale", 
                    Duplex = false, 
                    PaperSize = "A4",
                    UseADF = true
                }
            };
        }

        /// <summary>
        /// Get SharePoint document libraries
        /// </summary>
        private List<SelectListItem> GetTargetLibraries()
        {
            try
            {
                var libs = _helper.GetDocumentLibraries();
                return libs.Select(l => new SelectListItem { Text = l.Title, Value = l.Title }).ToList();
            }
            catch
            {
                return new List<SelectListItem>();
            }
        }

        /// <summary>
        /// Get classification options
        /// </summary>
        private List<SelectListItem> GetClassificationOptions()
        {
            return new List<SelectListItem>
            {
                new SelectListItem { Text = "Public", Value = "PUBLIC" },
                new SelectListItem { Text = "Internal", Value = "INTERNAL" },
                new SelectListItem { Text = "Confidential", Value = "CONFIDENTIAL" },
                new SelectListItem { Text = "Restricted", Value = "RESTRICTED" }
            };
        }

        /// <summary>
        /// Get document type options
        /// </summary>
        private List<SelectListItem> GetDocumentTypeOptions()
        {
            return new List<SelectListItem>
            {
                new SelectListItem { Text = "Invoice", Value = "Invoice" },
                new SelectListItem { Text = "Purchase Order", Value = "PurchaseOrder" },
                new SelectListItem { Text = "Contract", Value = "Contract" },
                new SelectListItem { Text = "Report", Value = "Report" },
                new SelectListItem { Text = "Memo", Value = "Memo" },
                new SelectListItem { Text = "Receipt", Value = "Receipt" }
            };
        }

        /// <summary>
        /// Get department options
        /// </summary>
        private List<SelectListItem> GetDepartmentOptions()
        {
            return new List<SelectListItem>
            {
                new SelectListItem { Text = "Finance", Value = "Finance" },
                new SelectListItem { Text = "Operations", Value = "Operations" },
                new SelectListItem { Text = "HR", Value = "HR" },
                new SelectListItem { Text = "IT", Value = "IT" },
                new SelectListItem { Text = "Sales", Value = "Sales" }
            };
        }

        /// <summary>
        /// Get available workflows
        /// </summary>
        private List<SelectListItem> GetWorkflowOptions()
        {
            try
            {
                var workflows = _helper.GetWorkflows();
                return workflows.Select(w => new SelectListItem { Text = w.WorkflowName, Value = w.Id.ToString() }).ToList();
            }
            catch
            {
                return new List<SelectListItem>();
            }
        }

        /// <summary>
        /// Generate unique document filename from metadata
        /// </summary>
        private string GenerateDocumentFileName(MetadataViewModel metadata)
        {
            return $"{metadata.ReferenceNumber}_{metadata.DocumentType}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
        }

        #endregion
    }
}