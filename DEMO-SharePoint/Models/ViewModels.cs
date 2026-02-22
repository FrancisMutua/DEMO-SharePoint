using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;
using DEMO_SharePoint.Models;

namespace ScanPortal.Web.ViewModels
{
    public class ScanSettingsViewModel
    {
        public List<string> AvailableScanners { get; set; } = new List<string>();
        public List<ScanProfile> SavedProfiles { get; set; } = new List<ScanProfile>();
        public List<SelectListItem> LibraryList { get; set; } = new List<SelectListItem>();
        [Required(ErrorMessage = "Please select a scanner")]
        public string SelectedScanner { get; set; }
        public int DPI { get; set; } = 300;
        public bool Duplex { get; set; }
        public string ColorMode { get; set; } = "Grayscale";
        public string PaperSize { get; set; } = "A4";
        public bool UseADF { get; set; } = true;
        public string SelectedProfile { get; set; }
        public string TargetLibrary { get; set; }
    }

    public class PreviewViewModel
    {
        public string FileBase64 { get; set; }
        public int PageCount { get; set; }
        public List<string> PagePreviews { get; set; } = new List<string>();
        public string OCRText { get; set; }
        public string BarcodeValue { get; set; }
        public string BarcodeFormat { get; set; }
    }

    public class MetadataViewModel
    {
        [Required(ErrorMessage = "Reference number is required")]
        [Display(Name = "Reference Number")]
        public string ReferenceNumber { get; set; }
        [Required(ErrorMessage = "Document type is required")]
        [Display(Name = "Document Type")]
        public string DocumentType { get; set; }
        [Required(ErrorMessage = "Classification code is required")]
        [Display(Name = "Classification Code")]
        public string ClassificationCode { get; set; }
        [Required(ErrorMessage = "Department is required")]
        [Display(Name = "Department")]
        public string Department { get; set; }
        [Required(ErrorMessage = "Record date is required")]
        [Display(Name = "Record Date")]
        [DataType(DataType.Date)]
        public DateTime RecordDate { get; set; } = DateTime.Today;
        [Display(Name = "Description / Subject")]
        public string Description { get; set; }
        [Display(Name = "Author / Originator")]
        public string Author { get; set; }
        [Required]
        public string TargetLibrary { get; set; }
        public string ContentType { get; set; }
        public bool SendToWorkflow { get; set; }
        public string WorkflowName { get; set; }
        public bool ArchiveImmediately { get; set; }
        public string OCRText { get; set; }
        public string BarcodeValue { get; set; }
        public string FileBase64 { get; set; }
        public int PageCount { get; set; }
        public List<MetadataField> DynamicFields { get; set; } = new List<MetadataField>();
        public Dictionary<string, string> DynamicFieldValues { get; set; } = new Dictionary<string, string>();
        public List<SelectListItem> ClassificationOptions { get; set; } = new List<SelectListItem>();
        public List<SelectListItem> DocumentTypeOptions { get; set; } = new List<SelectListItem>();
        public List<SelectListItem> DepartmentOptions { get; set; } = new List<SelectListItem>();
        public List<SelectListItem> LibraryOptions { get; set; } = new List<SelectListItem>();
        public List<SelectListItem> WorkflowOptions { get; set; } = new List<SelectListItem>();
    }

    public class ProfilesViewModel
    {
        public List<ScanProfile> Profiles { get; set; } = new List<ScanProfile>();
        public ScanProfile NewProfile { get; set; } = new ScanProfile();
        public List<SelectListItem> LibraryOptions { get; set; } = new List<SelectListItem>();
    }

    public class FieldsAdminViewModel
    {
        public List<MetadataField> ExistingFields { get; set; } = new List<MetadataField>();
        [Required]
        [Display(Name = "Field Display Name")]
        public string NewFieldDisplayName { get; set; }
        [Required]
        [Display(Name = "Field Type")]
        public string NewFieldType { get; set; }
        [Display(Name = "Required?")]
        public bool NewFieldIsRequired { get; set; }
        [Display(Name = "Choices (comma-separated, for Choice type)")]
        public string NewFieldChoices { get; set; }
        public string TargetLibrary { get; set; }
        public List<SelectListItem> LibraryOptions { get; set; } = new List<SelectListItem>();
    }

    public class UploadSuccessViewModel
    {
        public int ItemId { get; set; }
        public string FileUrl { get; set; }
        public string FileName { get; set; }
        public string Library { get; set; }
        public bool WorkflowStarted { get; set; }
        public bool Archived { get; set; }
    }

    /// <summary>
    /// Step 1 + 2: Main Scan Index Page ViewModel
    /// Used by ScanController.Index() to populate the scanner selection + settings UI
    /// </summary>
    public class ScanIndexViewModel
    {
        // ── Step 1: Scanner Selection ─────────────

        /// <summary>
        /// List of TWAIN scanner names retrieved from the local ScanPortal Agent
        /// running on the workstation. Populated via INetworkPrinterService.DiscoverNetworkScannersAsync().
        /// Empty if agent is offline.
        /// </summary>
        public List<string> AvailableScanners { get; set; }
            = new List<string>();

        /// <summary>
        /// The scanner name the user has selected.
        /// Passed to ScanController.InitiateScan() as part of ScanSettingsViewModel.
        /// </summary>
        [Display(Name = "Scanner")]
        public string SelectedScanner { get; set; }

        // ── Step 2: Scan Settings ─────────────────

        /// <summary>
        /// Scan resolution in dots per inch.
        /// Typical values: 150 (draft), 200 (standard), 300 (archive quality).
        /// </summary>
        [Display(Name = "Resolution (DPI)")]
        [Range(72, 1200, ErrorMessage = "DPI must be between 72 and 1200")]
        public int DPI { get; set; } = 300;

        /// <summary>
        /// Enable duplex (two-sided) scanning via the scanner's ADF unit.
        /// Requires scanner hardware support for duplex.
        /// </summary>
        [Display(Name = "Duplex (Two-Sided)")]
        public bool Duplex { get; set; } = false;

        /// <summary>
        /// Pixel/colour mode passed to TWAIN capability ICapPixelType.
        /// Accepted values: "Color" | "Grayscale" | "BlackWhite"
        /// </summary>
        [Display(Name = "Colour Mode")]
        public string ColorMode { get; set; } = "Grayscale";

        /// <summary>
        /// Paper size hint passed to the scanner.
        /// Accepted values: "A4" | "A3" | "Letter" | "Legal" | "Auto"
        /// </summary>
        [Display(Name = "Paper Size")]
        public string PaperSize { get; set; } = "A4";

        /// <summary>
        /// Use the Automatic Document Feeder (ADF) rather than flatbed.
        /// When true, enables multi-page continuous scanning.
        /// </summary>
        [Display(Name = "Use Auto Document Feeder (ADF)")]
        public bool UseADF { get; set; } = true;

        // ── Scan Profiles ─────────────────────────

        /// <summary>
        /// All saved scan profiles loaded from SharePoint 'ScanProfiles' custom list
        /// via INetworkPrinterService. Displayed in the "Load Profile" dropdown on the
        /// scan settings panel. Selecting a profile auto-fills DPI, Duplex, ColorMode,
        /// PaperSize, UseADF and TargetLibrary from the saved profile values.
        /// </summary>
        public List<ScanProfile> ScanProfiles { get; set; }
            = new List<ScanProfile>();

        /// <summary>
        /// ID of the currently selected profile (used to pre-select the dropdown).
        /// Null when no profile has been loaded.
        /// </summary>
        public int? SelectedProfileId { get; set; }

        // ── Target Library ────────────────────────

        /// <summary>
        /// SharePoint document library into which the scanned document
        /// will be uploaded. Pre-selected from the chosen scan profile's
        /// DefaultLibrary value, or overridden manually by the user.
        /// </summary>
        [Display(Name = "Target Library")]
        public string TargetLibrary { get; set; }

        /// <summary>
        /// All SharePoint document libraries available for selection,
        /// loaded via ISharePointService.GetLibrariesAsync().
        /// </summary>
        public List<SelectListItem> LibraryOptions { get; set; }
            = new List<SelectListItem>();

        // ── Agent Status ──────────────────────────

        /// <summary>
        /// True when the local scanner agent responded to the health check.
        /// False means the agent is not running on this workstation —
        /// the view renders an "Agent offline" warning banner.
        /// </summary>
        public bool AgentOnline { get; set; } = false;

        /// <summary>
        /// Error message returned when the agent health check or scanner
        /// enumeration fails. Displayed in the UI as a dismissible warning.
        /// Null when agent is online and scanners loaded successfully.
        /// </summary>
        public string AgentErrorMessage { get; set; }

        // ── Convenience Helpers ───────────────────

        /// <summary>
        /// True when at least one scanner is available from the agent.
        /// Used by the view to enable/disable the "Start Scan" button.
        /// </summary>
        public bool HasScanners =>
            AvailableScanners != null && AvailableScanners.Count > 0;

        /// <summary>
        /// True when at least one scan profile exists.
        /// Used by the view to show or hide the profile dropdown section.
        /// </summary>
        public bool HasProfiles =>
            ScanProfiles != null && ScanProfiles.Count > 0;

        /// <summary>
        /// Returns a SelectList built from ScanProfiles for use in the
        /// Razor Html.DropDownList helpers.
        /// </summary>
        public List<SelectListItem> ProfileSelectList =>
            ScanProfiles?.ConvertAll(p => new SelectListItem
            {
                Value = p.Id.ToString(),
                Text = $"{p.ProfileName}  ({p.DPI} DPI"
                       + (p.Duplex ? ", Duplex" : "")
                       + $", {p.ColorMode})",
                Selected = p.Id == SelectedProfileId
            }) ?? new List<SelectListItem>();

        /// <summary>
        /// Returns the currently selected ScanProfile object, or null
        /// if no profile is selected or SelectedProfileId is not found.
        /// </summary>
        public ScanProfile ActiveProfile =>
            SelectedProfileId.HasValue
                ? ScanProfiles?.Find(p => p.Id == SelectedProfileId.Value)
                : null;
    }
}
