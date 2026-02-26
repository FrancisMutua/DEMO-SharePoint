using DEMO_SharePoint.Models;
using KCAU_SharePoint.Models;
using Microsoft.SharePoint.ApplicationPages.ClientPickerQuery;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace DEMO_SharePoint.Controllers
{
    [SessionAuthorize]
    public class HomeController : Controller
    {

        Helper helper;

        public HomeController()
        {
            helper = new Helper();
        }

        public ActionResult Index()
        {
            var libraries = helper.GetDocumentLibraries();
            ViewBag.Libraries = libraries;
            var model = new DashboardViewModel
            {
                RecentActivities = new List<RecentActivity>()
            };

            using (var context = helper.GetContext())
            {

                Web web = context.Web;
                ListCollection lists = web.Lists;
                context.Load(lists, l => l.Include(
                    list => list.Title,
                    list => list.BaseTemplate,
                    list => list.ItemCount,
                    list => list.Hidden
                ));
                context.ExecuteQuery();

                // Total documents
                model.TotalDocuments = lists.Where(l => l.BaseTemplate == 101 && !l.Hidden).Sum(l => l.ItemCount);

                // Recent activities (last 7 days)
                foreach (var list in lists.Where(l => l.BaseTemplate == 101 && !l.Hidden))
                {
                    CamlQuery query = new CamlQuery
                    {
                        ViewXml = @"<View>
                                <Query>
                                    <Where>
                                        <Geq>
                                            <FieldRef Name='Created'/>
                                            <Value Type='DateTime'>
                                                <Today OffsetDays='-7'/>
                                            </Value>
                                        </Geq>
                                    </Where>
                                </Query>
                                <RowLimit>10</RowLimit>
                            </View>"
                    };

                    ListItemCollection items = list.GetItems(query);
                    context.Load(items, i => i.Include(
                        it => it["Author"],
                        it => it["Editor"],
                        it => it["FileLeafRef"],
                        it => it["Created"]
                    ));
                    context.ExecuteQuery();

                    foreach (var item in items)
                    {
                        model.RecentActivities.Add(new RecentActivity
                        {
                            User = ((FieldUserValue)item["Author"]).LookupValue,
                            Activity = "Uploaded Document",
                            Document = item["FileLeafRef"]?.ToString(),
                            Date = Convert.ToDateTime(item["Created"])
                        });
                    }
                }

                // Uploads today
                model.UploadsToday = model.RecentActivities.Count(a => a.Date.Date == DateTime.Now.Date);

                // Active users
                model.ActiveUsers = model.RecentActivities.Select(a => a.User).Distinct().Count();
            }

            return View(model);

        }
        public ActionResult DocumentLibrary(string libraryUrl)
        {
            var libraries = helper.GetDocumentLibraries();
            ViewBag.Libraries = libraries;

            if (string.IsNullOrEmpty(libraryUrl))
                return RedirectToAction("Index");

            ViewBag.LibraryUrl = libraryUrl;

            // Get all files and folders
            var items = GetFolderContents(libraryUrl);
            ViewBag.Items = items;

            return View();
        }

        /// <summary>
        /// Creates a new document library
        /// </summary>
        [HttpPost]
        [ValidateInput(false)]
        public JsonResult CreateLibrary(string libraryName, string description)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(libraryName))
                    return Json(new { success = false, message = "Library name is required" });

                using (var ctx = helper.GetContext())
                {
                    Web web = ctx.Web;
                    ListCreationInformation createInfo = new ListCreationInformation
                    {
                        Title = libraryName,
                        Description = description ?? "",
                        TemplateType = (int)ListTemplateType.DocumentLibrary
                    };

                    List newList = web.Lists.Add(createInfo);
                    ctx.ExecuteQuery();

                    // Enable versioning on the new library
                    newList.EnableVersioning = true;
                    newList.MajorVersionLimit = 100;
                    newList.Update();
                    ctx.ExecuteQuery();

                    return Json(new
                    {
                        success = true,
                        libraryTitle = newList.Title,
                        libraryUrl = newList.RootFolder.ServerRelativeUrl
                    });
                }
            }
            catch (ServerException ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Updates library metadata (name, description)
        /// </summary>
        [HttpPost]
        [ValidateInput(false)]
        public JsonResult UpdateLibraryMetadata(string libraryUrl, string newTitle, string newDescription)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(libraryUrl))
                    return Json(new { success = false, message = "Library URL is required" });

                if (string.IsNullOrWhiteSpace(newTitle))
                    return Json(new { success = false, message = "Library name is required" });

                using (var ctx = helper.GetContext())
                {
                    Web web = ctx.Web;
                    List list = web.GetListByTitle(ExtractLibraryTitle(libraryUrl));

                    ctx.Load(list);
                    ctx.ExecuteQuery();

                    list.Title = newTitle;
                    list.Description = newDescription ?? "";
                    list.Update();
                    ctx.ExecuteQuery();

                    return Json(new { success = true, message = "Library updated successfully" });
                }
            }
            catch (ServerException ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Gets library details for editing
        /// </summary>
        [HttpGet]
        public JsonResult GetLibraryDetails(string libraryUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(libraryUrl))
                    return Json(new { success = false, message = "Library URL is required" }, JsonRequestBehavior.AllowGet);

                using (var ctx = helper.GetContext())
                {
                    Web web = ctx.Web;

                    var folder = web.GetFolderByServerRelativeUrl(libraryUrl);
                    ctx.Load(folder);
                    ctx.ExecuteQuery();

                    var lists = web.Lists;
                    ctx.Load(lists, l => l.Include(
                        list => list.Title,
                        list => list.Description,
                        list => list.RootFolder.ServerRelativeUrl,
                        list => list.ItemCount,
                        list => list.EnableVersioning
                    ));
                    ctx.ExecuteQuery();

                    var library = lists.FirstOrDefault(l => l.RootFolder.ServerRelativeUrl.Equals(libraryUrl, StringComparison.OrdinalIgnoreCase));

                    if (library == null)
                        return Json(new { success = false, message = "Library not found" }, JsonRequestBehavior.AllowGet);

                    return Json(new
                    {
                        success = true,
                        title = library.Title,
                        description = library.Description ?? "",
                        itemCount = library.ItemCount,
                        versioningEnabled = library.EnableVersioning,
                        url = library.RootFolder.ServerRelativeUrl
                    }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// Deletes a document library from SharePoint
        /// </summary>
        [HttpPost]
        public JsonResult DeleteLibrary(string libraryUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(libraryUrl))
                    return Json(new { success = false, message = "Library URL is required" });

                using (var ctx = helper.GetContext())
                {
                    Web web = ctx.Web;
                    var lists = web.Lists;
                    ctx.Load(lists);
                    ctx.ExecuteQuery();

                    var library = lists.FirstOrDefault(l => l.RootFolder.ServerRelativeUrl.Equals(libraryUrl, StringComparison.OrdinalIgnoreCase));

                    if (library == null)
                        return Json(new { success = false, message = "Library not found" });

                    library.DeleteObject();
                    ctx.ExecuteQuery();

                    return Json(new { success = true, message = "Library deleted successfully" });
                }
            }
            catch (ServerException ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Creates a new folder in a document library
        /// </summary>
        [HttpPost]
        [ValidateInput(false)]
        public JsonResult CreateFolder(string folderName, string libraryUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(folderName))
                    return Json(new { success = false, message = "Folder name missing" });

                if (string.IsNullOrWhiteSpace(libraryUrl))
                    return Json(new { success = false, message = "Library URL missing" });

                using (var ctx = helper.GetContext())
                {

                    // 🔑 MUST be server-relative
                    var folder = ctx.Web.GetFolderByServerRelativeUrl(libraryUrl);

                    ctx.Load(folder);
                    ctx.ExecuteQuery();

                    folder.Folders.Add(folderName);
                    ctx.ExecuteQuery();
                }

                return Json(new { success = true });
            }
            catch (ServerException ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public JsonResult SubmitForApproval(string itemUrl, string itemName)
        {
            try
            {
                var workflow = helper.GetWorkflowForLibrary(helper.GetLibraryUrlFromItem(itemUrl));

                if (workflow == null)
                    return Json(new { success = false, message = "No workflow configured." });

                // Replace hard-coded user for demo; in real app, use User.Identity.Name
                helper.CreateWorkflowInstance(itemUrl, itemName, User.Identity.Name);

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public JsonResult RenameItem(string itemUrl, bool isFolder, string newName)
        {
            try
            {
                using (var ctx = helper.GetContext())
                {


                    // Ensure server-relative URL
                    if (itemUrl.StartsWith("http"))
                        itemUrl = new Uri(itemUrl).AbsolutePath;

                    if (isFolder)
                    {
                        var folder = ctx.Web.GetFolderByServerRelativeUrl(itemUrl);
                        ctx.Load(folder, f => f.ServerRelativeUrl);
                        ctx.ExecuteQuery();

                        string parentUrl = itemUrl.Substring(0, itemUrl.LastIndexOf('/'));
                        string targetUrl = parentUrl + "/" + newName;

                        folder.MoveTo(targetUrl);
                    }
                    else
                    {
                        var file = ctx.Web.GetFileByServerRelativeUrl(itemUrl);
                        ctx.Load(file, f => f.ServerRelativeUrl);
                        ctx.ExecuteQuery();

                        // Keep file extension
                        string ext = Path.GetExtension(file.ServerRelativeUrl);
                        if (!newName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                            newName += ext;

                        string parentUrl = itemUrl.Substring(0, itemUrl.LastIndexOf('/'));
                        string targetUrl = parentUrl + "/" + newName;

                        file.MoveTo(targetUrl, MoveOperations.Overwrite);
                    }

                    ctx.ExecuteQuery();
                }

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public JsonResult DeleteItem(string itemUrl, bool isFolder)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(itemUrl))
                    return Json(new { success = false, message = "Item URL missing" });

                using (var ctx = helper.GetContext())
                {

                    if (isFolder)
                    {
                        var folder = ctx.Web.GetFolderByServerRelativeUrl(itemUrl);
                        folder.DeleteObject();
                    }
                    else
                    {
                        var file = ctx.Web.GetFileByServerRelativeUrl(itemUrl);
                        file.DeleteObject();
                    }

                    ctx.ExecuteQuery();
                }

                return Json(new { success = true });
            }
            catch (ServerException ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public JsonResult UploadFile(IEnumerable<HttpPostedFileBase> files, string libraryUrl)
        {
            try
            {
                if (files == null || !files.Any())
                    return Json(new { success = false, message = "No files selected" });

                if (string.IsNullOrWhiteSpace(libraryUrl))
                    return Json(new { success = false, message = "Library URL missing" });

                var uploadResults = new List<object>();

                using (var ctx = helper.GetContext())
                {
                    var folder = ctx.Web.GetFolderByServerRelativeUrl(libraryUrl);
                    ctx.Load(folder);
                    ctx.ExecuteQuery();

                    // Process each file
                    foreach (var file in files)
                    {
                        if (file == null || file.ContentLength == 0)
                        {
                            uploadResults.Add(new
                            {
                                fileName = "Unknown",
                                success = false,
                                message = "File is empty or invalid"
                            });
                            continue;
                        }

                        try
                        {
                            var fileInfo = new FileCreationInformation
                            {
                                ContentStream = file.InputStream,
                                Url = file.FileName,
                                Overwrite = true
                            };

                            var upload = folder.Files.Add(fileInfo);
                            ctx.Load(upload);
                            ctx.ExecuteQuery();

                            uploadResults.Add(new
                            {
                                fileName = file.FileName,
                                success = true,
                                message = "Uploaded successfully"
                            });
                        }
                        catch (Exception ex)
                        {
                            uploadResults.Add(new
                            {
                                fileName = file.FileName,
                                success = false,
                                message = ex.Message
                            });
                        }
                    }
                }

                // Return summary of all uploads
                bool allSuccess = uploadResults.All(r => (bool)((dynamic)r).success);
                return Json(new
                {
                    success = allSuccess,
                    message = allSuccess ? "All files uploaded successfully" : "Some files failed to upload",
                    uploadedFiles = uploadResults,
                    totalFiles = uploadResults.Count,
                    successCount = uploadResults.Count(r => (bool)((dynamic)r).success),
                    failureCount = uploadResults.Count(r => !(bool)((dynamic)r).success)
                });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Upload error: {ex.Message}" });
            }
        }


        private static readonly HashSet<string> ModernViewFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    "FileLeafRef",       // Name
    "Modified",          // Modified date
    "Editor",            // Modified By
    "File_x0020_Size",   // File Size
    "FileSizeDisplay",   // File Size (display)
    "FSObjType",         // Is Folder flag
    "UniqueId",          // Unique identifier
    "ContentTypeId",     // Content Type
    "ContentType"        // Content Type name
};

        private List<SPItem> GetFolderContents(string folderUrl)
        {
            var items = new List<SPItem>();
            using (var ctx = helper.GetContext())
            {
                // Get library root URL to fetch the list
                // e.g. /sites/MySite/Shared Documents/SubFolder -> /sites/MySite/Shared Documents
                var listRootUrl = GetListRootUrl(folderUrl);
                var list = ctx.Web.GetList(listRootUrl);
                var view = list.DefaultView;
                ctx.Load(view, v => v.ViewFields);
                ctx.ExecuteQuery();

                var viewFields = new HashSet<string>(view.ViewFields, StringComparer.OrdinalIgnoreCase);

                var folder = ctx.Web.GetFolderByServerRelativeUrl(folderUrl);
                ctx.Load(folder);
                ctx.Load(folder.Folders, f => f.Include(
                    ff => ff.Name,
                    ff => ff.ServerRelativeUrl,
                    ff => ff.ListItemAllFields));
                ctx.Load(folder.Files, f => f.Include(
                    ff => ff.Name,
                    ff => ff.ServerRelativeUrl,
                    ff => ff.Length,
                    ff => ff.ListItemAllFields));
                ctx.ExecuteQuery();

                // Folders
                foreach (var subFolder in folder.Folders)
                {
                    if (!subFolder.Name.StartsWith("Forms"))
                    {
                        var spItem = new SPItem
                        {
                            Name = subFolder.Name,
                            Url = subFolder.ServerRelativeUrl,
                            IsFolder = true
                        };
                        ExtractViewMetadata(subFolder.ListItemAllFields, spItem, viewFields);
                        items.Add(spItem);
                    }
                }

                // Files
                foreach (var file in folder.Files)
                {
                    var spItem = new SPItem
                    {
                        Name = file.Name,
                        Url = file.ServerRelativeUrl,
                        IsFolder = false
                    };
                    ExtractViewMetadata(file.ListItemAllFields, spItem, viewFields);
                    items.Add(spItem);
                }
            }
            return items;
        }

        /// <summary>
        /// Walks up the folderUrl until it finds the list root.
        /// e.g. /sites/MySite/Shared Documents/Folder1/SubFolder -> /sites/MySite/Shared Documents
        /// </summary>
        private string GetListRootUrl(string folderUrl)
        {
            using (var ctx = helper.GetContext())
            {
                // Try each segment from full path upward until GetList succeeds
                var parts = folderUrl.TrimEnd('/').Split('/');
                for (int i = parts.Length; i >= 2; i--)
                {
                    var candidateUrl = string.Join("/", parts.Take(i));
                    try
                    {
                        var list = ctx.Web.GetList(candidateUrl);
                        ctx.Load(list, l => l.Title);
                        ctx.ExecuteQuery();
                        return candidateUrl; // Found the list root
                    }
                    catch (ServerException)
                    {
                        // Not a list at this level, go one level up
                    }
                }
            }
            throw new Exception($"Could not find list root for folder: {folderUrl}");
        }

        private void ExtractViewMetadata(ListItem listItem, SPItem spItem, HashSet<string> viewFields)
        {
            if (listItem == null) return;

            foreach (var field in listItem.FieldValues)
            {
                if (viewFields.Contains(field.Key))
                {
                    if (field.Value != null)
                        spItem.Metadata[field.Key] = ConvertFieldValue(field.Value);
                    else
                        spItem.Metadata[field.Key] = "";
                }
            }
        }

        private string ConvertFieldValue(object value)
        {
            switch (value)
            {
                // Modified By, Created By
                case FieldUserValue userValue:
                    return userValue.LookupValue; // Display name e.g. "John Smith"

                // Multiple users
                case FieldUserValue[] userValues:
                    return string.Join(", ", userValues.Select(u => u.LookupValue));

                // Lookup field
                case FieldLookupValue lookupValue:
                    return lookupValue.LookupValue;

                // Multiple lookups
                case FieldLookupValue[] lookupValues:
                    return string.Join(", ", lookupValues.Select(l => l.LookupValue));

                // Managed metadata / taxonomy
                case TaxonomyFieldValue taxonomyValue:
                    return taxonomyValue.Label;

                // Multiple managed metadata
                case TaxonomyFieldValueCollection taxonomyValues:
                    return string.Join(", ", taxonomyValues.Select(t => t.Label));

                // URL field
                case FieldUrlValue urlValue:
                    return urlValue.Url;

                // DateTime
                case DateTime dt:
                    return dt.ToLocalTime().ToString("g"); // e.g. 2/22/2026 10:30 AM

                // Boolean
                case bool b:
                    return b ? "Yes" : "No";

                // Default
                default:
                    return value.ToString();
            }
        }
        /// <summary>
        /// Returns file content as binary blob with proper MIME type headers.
        /// Supports PDFs, Office documents, images, text files, and more.
        /// Client-side JavaScript handles rendering based on file type.
        /// </summary>
        [HttpPost]
        public ActionResult ViewDocument(string fileUrl)
        {
            if (string.IsNullOrWhiteSpace(fileUrl))
                return new HttpStatusCodeResult(400, "Invalid file url");

            try
            {
                using (var ctx = helper.GetContext())
                {
                    var file = ctx.Web.GetFileByServerRelativeUrl(fileUrl);
                    ctx.Load(file);
                    ctx.ExecuteQuery();

                    var stream = file.OpenBinaryStream();
                    ctx.ExecuteQuery();

                    using (var ms = new MemoryStream())
                    {
                        stream.Value.CopyTo(ms);
                        var bytes = ms.ToArray();
                        var fileName = file.Name;
                        var extension = Path.GetExtension(fileName).ToLower();

                        // Determine MIME type based on extension
                        var mimeType = GetMimeType(extension);

                        // Set response headers for inline viewing
                        Response.AddHeader("Content-Disposition", $"inline; filename=\"{fileName}\"");
                        Response.Cache.SetCacheability(HttpCacheability.Public);
                        Response.Cache.SetMaxAge(TimeSpan.FromHours(1));
                        Response.AddHeader("Access-Control-Allow-Origin", "*");

                        // Return file as binary blob
                        return File(bytes, mimeType, fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ViewDocument error: {ex.Message}");
                // Return error as JSON for client-side error handling
                return Json(new
                {
                    success = false,
                    message = $"Error loading document: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// Maps file extensions to their correct MIME types.
        /// Comprehensive support for documents, images, archives, and more.
        /// </summary>
        private string GetMimeType(string extension)
        {
            var mimeTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Microsoft Office Documents (Legacy & Modern)
                { ".doc", "application/msword" },
                { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                { ".docm", "application/vnd.ms-word.document.macroEnabled.12" },
                { ".dot", "application/msword" },
                { ".dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template" },

                { ".xls", "application/vnd.ms-excel" },
                { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                { ".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12" },
                { ".xlt", "application/vnd.ms-excel" },
                { ".xltx", "application/vnd.openxmlformats-officedocument.spreadsheetml.template" },

                { ".ppt", "application/vnd.ms-powerpoint" },
                { ".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
                { ".pptm", "application/vnd.ms-powerpoint.presentation.macroEnabled.12" },
                { ".pot", "application/vnd.ms-powerpoint" },
                { ".potx", "application/vnd.openxmlformats-officedocument.presentationml.template" },
                { ".ppsx", "application/vnd.openxmlformats-officedocument.presentationml.slideshow" },
                
                // OpenOffice/LibreOffice Documents
                { ".odt", "application/vnd.oasis.opendocument.text" },
                { ".ods", "application/vnd.oasis.opendocument.spreadsheet" },
                { ".odp", "application/vnd.oasis.opendocument.presentation" },

                // Adobe & PDF
                { ".pdf", "application/pdf" },
                { ".ai", "application/postscript" },
                { ".eps", "application/postscript" },

                // Text Documents
                { ".txt", "text/plain" },
                { ".rtf", "application/rtf" },
                { ".csv", "text/csv" },
                { ".tsv", "text/tab-separated-values" },
                { ".log", "text/plain" },
                { ".md", "text/markdown" },
                { ".tex", "text/x-tex" },

                // Markup & Data
                { ".xml", "application/xml" },
                { ".json", "application/json" },
                { ".html", "text/html" },
                { ".htm", "text/html" },
                { ".xhtml", "application/xhtml+xml" },
                { ".svg", "image/svg+xml" },

                // Images - Raster
                { ".jpg", "image/jpeg" },
                { ".jpeg", "image/jpeg" },
                { ".jpe", "image/jpeg" },
                { ".png", "image/png" },
                { ".gif", "image/gif" },
                { ".bmp", "image/bmp" },
                { ".dib", "image/bmp" },
                { ".webp", "image/webp" },
                { ".tiff", "image/tiff" },
                { ".tif", "image/tiff" },
                { ".ico", "image/x-icon" },
                { ".cur", "image/x-icon" },

                // Images - Vector
                { ".emf", "application/x-msmetafile" },
                { ".wmf", "application/x-msmetafile" },

                // Archives & Compressed
                { ".zip", "application/zip" },
                { ".rar", "application/x-rar-compressed" },
                { ".7z", "application/x-7z-compressed" },
                { ".tar", "application/x-tar" },
                { ".gz", "application/gzip" },
                { ".tgz", "application/x-tar+gzip" },
                { ".bz2", "application/x-bzip2" },

                // Audio
                { ".mp3", "audio/mpeg" },
                { ".wav", "audio/wav" },
                { ".flac", "audio/flac" },
                { ".aac", "audio/aac" },
                { ".ogg", "audio/ogg" },
                { ".m4a", "audio/mp4" },

                // Video
                { ".mp4", "video/mp4" },
                { ".avi", "video/x-msvideo" },
                { ".mov", "video/quicktime" },
                { ".mkv", "video/x-matroska" },
                { ".webm", "video/webm" },
                { ".flv", "video/x-flv" },
                { ".wmv", "video/x-ms-wmv" },

                // Presentations
                { ".vsd", "application/vnd.visio" },
                { ".vsdx", "application/vnd.ms-visio.drawing" },
                { ".mpp", "application/vnd.ms-project" }
            };

            // Return mapped MIME type or generic binary if extension not found
            return mimeTypes.ContainsKey(extension) ? mimeTypes[extension] : "application/octet-stream";
        }

        [HttpPost]
        public JsonResult GetFileVersions(string fileUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(fileUrl))
                    return Json(new { success = false, message = "Invalid file URL" });

                using (var ctx = helper.GetContext())
                {

                    var file = ctx.Web.GetFileByServerRelativeUrl(fileUrl);
                    ctx.Load(file, f => f.Versions);
                    ctx.ExecuteQuery();

                    var versions = file.Versions.Select(v => new
                    {
                        Url = file.ServerRelativeUrl,
                        VersionLabel = v.VersionLabel,
                        Created = v.Created
                        // CreatedBy = v.CreatedBy?.Title
                    }).OrderByDescending(v => v.VersionLabel).ToList();

                    return Json(new { success = true, versions });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public JsonResult GrantItemAccess(string itemUrl, string userLogin, string role)
        {
            try
            {
                using (var ctx = helper.GetContext())
                {

                    // Load web roles
                    ctx.Load(ctx.Web, w => w.RoleDefinitions, w => w.Url);
                    ctx.ExecuteQuery();

                    // Get item
                    var file = ctx.Web.GetFileByServerRelativeUrl(itemUrl);
                    var item = file.ListItemAllFields;

                    ctx.Load(item, i => i.HasUniqueRoleAssignments);
                    ctx.ExecuteQuery();

                    // Break inheritance if needed
                    if (!item.HasUniqueRoleAssignments)
                    {
                        item.BreakRoleInheritance(false, true);
                        ctx.ExecuteQuery();
                    }

                    // Ensure user
                    var spUser = ctx.Web.EnsureUser(userLogin);
                    ctx.Load(spUser, u => u.Email, u => u.Title);
                    ctx.ExecuteQuery();

                    // Assign role
                    var roleDef = ctx.Web.RoleDefinitions.GetByName(role);
                    var bindings = new RoleDefinitionBindingCollection(ctx) { roleDef };

                    item.RoleAssignments.Add(spUser, bindings);
                    ctx.ExecuteQuery();

                    // 🔔 SEND ALERT EMAIL
                    if (!string.IsNullOrWhiteSpace(spUser.Email))
                    {
                        SendShareNotification(
                            ctx,
                            spUser.Email,
                            spUser.Title,
                            ctx.Web.Url + itemUrl,
                            role
                        );
                    }
                }

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private void SendShareNotification(ClientContext ctx, string toEmail, string displayName, string itemLink, string role)
        {
            var emailProps = new EmailProperties
            {
                To = new[] { toEmail },
                Subject = "A document has been shared with you",
                Body = $@"
                            Hello {displayName},<br/><br/>
                            A document has been shared with you with <b>{role}</b> access.<br/><br/>
                            <a href='{itemLink}'>Click here to open the document</a><br/><br/>
                            Regards,<br/>
                            SharePoint System
                        "
            };

            Utility.SendEmail(ctx, emailProps);
            ctx.ExecuteQuery();
        }

        [HttpGet]
        public JsonResult SearchUsers(string term)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(term))
                    return Json(new List<object>(), JsonRequestBehavior.AllowGet);

                var results = new List<object>();

                using (var ctx = helper.GetContext())
                {
                    // Load all users in the site
                    var users = ctx.Web.SiteUsers;
                    ctx.Load(users, u => u.Include(
                        usr => usr.Title,
                        usr => usr.LoginName,
                        usr => usr.Email));
                    ctx.ExecuteQuery();

                    // Prepare login names for AD query (remove claims prefix)
                    var adLoginNames = users
                        .Where(u => string.IsNullOrWhiteSpace(u.Email))
                        .Select(u => CleanLoginName(u.LoginName))
                        .Where(u => !string.IsNullOrWhiteSpace(u))
                        .ToList();

                    // Query AD in batch
                    var adEmails = GetEmailsFromAD(adLoginNames);


                    // Merge SharePoint users with AD emails
                    foreach (var user in users)
                    {
                        string email = user.Email;

                        if (string.IsNullOrWhiteSpace(email))
                        {
                            string cleanLogin = CleanLoginName(user.LoginName);
                            if (cleanLogin != null && adEmails.ContainsKey(cleanLogin))
                                email = adEmails[cleanLogin];
                        }

                        // Filter by term
                        if ((user.Title != null && user.Title.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (user.LoginName != null && user.LoginName.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (email != null && email.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            results.Add(new
                            {
                                DisplayName = user.Title,
                                Login = NormalizeLogin(user.LoginName), // ✅ CLEAN LOGIN
                                Email = email
                            });

                        }
                    }
                }

                return Json(results, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new List<object>(), JsonRequestBehavior.AllowGet);
            }
        }

        private string NormalizeLogin(string loginName)
        {
            if (string.IsNullOrWhiteSpace(loginName))
                return null;

            // Remove claims if present
            if (loginName.StartsWith("i:0#.w|", StringComparison.OrdinalIgnoreCase))
                loginName = loginName.Substring("i:0#.w|".Length);

            // Remove domain if already present
            if (loginName.Contains("\\"))
                loginName = loginName.Split('\\').Last();

            return loginName;
        }

        // Batch query AD for multiple users
        private Dictionary<string, string> GetEmailsFromAD(List<string> loginNames)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var context = helper.GetPrincipalContext())
                {
                    using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                    {
                        foreach (var principal in searcher.FindAll().OfType<UserPrincipal>())
                        {
                            if (principal.SamAccountName != null && loginNames.Contains(principal.SamAccountName))
                            {
                                result[principal.SamAccountName] = principal.EmailAddress;
                            }
                        }
                    }
                }
            }
            catch
            {
                // AD not reachable
            }

            return result;
        }

        // Remove SharePoint claims prefix
        private string CleanLoginName(string loginName)
        {
            if (string.IsNullOrWhiteSpace(loginName))
                return null;

            if (loginName.Contains("|"))
                loginName = loginName.Split('|').Last();

            if (loginName.Contains("\\"))
                loginName = loginName.Split('\\').Last(); // get username only

            return loginName;
        }

        private string ExtractLibraryTitle(string libraryUrl)
        {
            if (string.IsNullOrWhiteSpace(libraryUrl))
                return null;

            // Remove leading slash if present
            string url = libraryUrl.TrimStart('/');

            // Get the last part of the URL path
            // Example: "sites/site/DocumentLibrary" → "DocumentLibrary"
            string[] parts = url.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

            return parts.Length > 0 ? parts[parts.Length - 1] : null;
        }

        /// <summary>
        /// Gets all custom columns defined on the library (for metadata form)
        /// </summary>
        [HttpGet]
        public JsonResult GetLibraryFields(string libraryUrl)
        {
            try
            {
                var listRootUrl = GetListRootUrl(libraryUrl);
                using (var ctx = helper.GetContext())
                {
                    var list = ctx.Web.GetList(listRootUrl);
                    ctx.Load(list.Fields, fields => fields.Include(
                        f => f.InternalName,
                        f => f.Title,
                        f => f.FieldTypeKind,
                        f => f.Required,
                        f => f.Hidden,
                        f => f.ReadOnlyField,
                        f => f.SchemaXml
                    ));
                    ctx.ExecuteQuery();

                    // Only return user-editable fields (same as modern view shows)
                    var editableFields = list.Fields
                        .Where(f =>
                            !f.Hidden &&
                            !f.ReadOnlyField &&
                            f.FieldTypeKind != FieldType.Computed &&
                            f.FieldTypeKind != FieldType.File &&
                            !f.InternalName.StartsWith("_") &&
                            !f.InternalName.StartsWith("vti_") &&
                            f.InternalName != "ContentType" &&
                            f.InternalName != "ContentTypeId" &&
                            f.InternalName != "FileLeafRef" // Name is handled separately
                        )
                        .Select(f => new
                        {
                            internalName = f.InternalName,
                            title = f.Title,
                            fieldType = f.FieldTypeKind.ToString(),
                            required = f.Required,
                            // For choice fields, extract choices from schema
                            choices = f.FieldTypeKind == FieldType.Choice
                                ? ExtractChoices(f.SchemaXml)
                                : null
                        })
                        .ToList();

                    return Json(new { success = true, fields = editableFields }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// Saves metadata on a file or folder item
        /// </summary>
        [HttpPost]
        public JsonResult SaveItemMetadata(string itemUrl, bool isFolder, Dictionary<string, string> metadata)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(itemUrl))
                    return Json(new { success = false, message = "Item URL missing" });

                using (var ctx = helper.GetContext())
                {
                    ListItem listItem;

                    if (isFolder)
                    {
                        var folder = ctx.Web.GetFolderByServerRelativeUrl(itemUrl);
                        ctx.Load(folder, f => f.ListItemAllFields);
                        ctx.ExecuteQuery();
                        listItem = folder.ListItemAllFields;
                    }
                    else
                    {
                        var file = ctx.Web.GetFileByServerRelativeUrl(itemUrl);
                        ctx.Load(file, f => f.ListItemAllFields);
                        ctx.ExecuteQuery();
                        listItem = file.ListItemAllFields;
                    }

                    if (metadata != null)
                    {
                        foreach (var kvp in metadata)
                        {
                            if (!string.IsNullOrWhiteSpace(kvp.Key))
                                listItem[kvp.Key] = kvp.Value;
                        }
                    }

                    listItem.Update();
                    ctx.ExecuteQuery();

                    return Json(new { success = true });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private List<string> ExtractChoices(string schemaXml)
        {
            var choices = new List<string>();
            try
            {
                var doc = System.Xml.Linq.XDocument.Parse(schemaXml);
                choices = doc.Descendants("CHOICE").Select(c => c.Value).ToList();
            }
            catch { }
            return choices;
        }

        /// <summary>
        /// Uploads a folder with its complete structure (nested folders and files)
        /// </summary>
        [HttpPost]
        [ValidateInput(false)]
        public JsonResult UploadFolder(string libraryUrl)
        {
            try
            {
                var folderItems = Request.Files.GetMultiple("folderItems");

                if (folderItems == null || folderItems.Count == 0)
                    return Json(new { success = false, message = "No folder items selected" });

                if (string.IsNullOrWhiteSpace(libraryUrl))
                    return Json(new { success = false, message = "Library URL missing" });

                var uploadResults = new List<object>();
                var createdFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                using (var ctx = helper.GetContext())
                {
                    var libraryFolder = ctx.Web.GetFolderByServerRelativeUrl(libraryUrl);
                    ctx.Load(libraryFolder);
                    ctx.ExecuteQuery();

                    // Get root folder name from first item's path
                    string rootFolderName = null;
                    if (folderItems.Count > 0 && !string.IsNullOrEmpty(folderItems[0].FileName))
                    {
                        var firstItemPath = folderItems[0].FileName;
                        rootFolderName = firstItemPath.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries)[0];
                    }

                    foreach (var fileItem in folderItems)
                    {
                        if (fileItem == null || fileItem.ContentLength == 0)
                            continue;

                        try
                        {
                            // Extract relative path from webkitRelativePath
                            string filePath = fileItem.FileName; // e.g., "FolderName/SubFolder/file.txt"
                            string[] pathParts = filePath.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);

                            if (pathParts.Length == 0)
                                continue;

                            // Build folder structure in SharePoint
                            string currentFolderUrl = libraryUrl;

                            for (int i = 0; i < pathParts.Length - 1; i++)
                            {
                                string folderName = pathParts[i];
                                string potentialFolderUrl = currentFolderUrl + "/" + folderName;

                                // Create folder if it doesn't exist
                                if (!createdFolders.Contains(potentialFolderUrl))
                                {
                                    try
                                    {
                                        var folder = ctx.Web.GetFolderByServerRelativeUrl(potentialFolderUrl);
                                        ctx.Load(folder);
                                        ctx.ExecuteQuery();
                                    }
                                    catch
                                    {
                                        // Folder doesn't exist, create it
                                        var parentFolder = ctx.Web.GetFolderByServerRelativeUrl(currentFolderUrl);
                                        parentFolder.Folders.Add(folderName);
                                        ctx.ExecuteQuery();
                                    }

                                    createdFolders.Add(potentialFolderUrl);
                                }

                                currentFolderUrl = potentialFolderUrl;
                            }

                            // Upload the file
                            string fileName = pathParts[pathParts.Length - 1];
                            var fileInfo = new FileCreationInformation
                            {
                                ContentStream = fileItem.InputStream,
                                Url = fileName,
                                Overwrite = true
                            };

                            var uploadFolder = ctx.Web.GetFolderByServerRelativeUrl(currentFolderUrl);
                            var uploadedFile = uploadFolder.Files.Add(fileInfo);
                            ctx.Load(uploadedFile);
                            ctx.ExecuteQuery();

                            uploadResults.Add(new
                            {
                                fileName = filePath,
                                success = true,
                                message = "Uploaded successfully"
                            });
                        }
                        catch (Exception ex)
                        {
                            uploadResults.Add(new
                            {
                                fileName = fileItem.FileName,
                                success = false,
                                message = ex.Message
                            });
                        }
                    }
                }

                bool allSuccess = uploadResults.All(r => (bool)((dynamic)r).success);
                return Json(new
                {
                    success = allSuccess,
                    message = allSuccess ? "Folder uploaded successfully" : "Some items failed to upload",
                    uploadedFiles = uploadResults,
                    totalFiles = uploadResults.Count,
                    successCount = uploadResults.Count(r => (bool)((dynamic)r).success),
                    failureCount = uploadResults.Count(r => !(bool)((dynamic)r).success)
                });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Upload error: {ex.Message}" });
            }
        }

        // ═══════════════════════════════════════════════
        //  LIBRARY MANAGER — Save & Delete
        // ═══════════════════════════════════════════════

        /// <summary>
        /// Creates or updates a document library and adds its metadata fields in SharePoint.
        /// Called from the Library Manager view (fire-and-forget on client, but fully implemented here).
        /// </summary>
        [HttpPost]
        public ActionResult SaveLibrary(LibraryDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.Name) || string.IsNullOrWhiteSpace(dto.Url))
                return Json(new { success = false, message = "Invalid library data." });

            try
            {
                using (var ctx = helper.GetContext())
                {
                    var web = ctx.Web;
                    List library = null;

                    // 1️⃣ CHECK IF LIBRARY EXISTS — try to load by title
                    try
                    {
                        library = web.Lists.GetByTitle(dto.Name);
                        ctx.Load(library, l => l.Title, l => l.RootFolder.ServerRelativeUrl);
                        ctx.ExecuteQuery();
                    }
                    catch
                    {
                        library = null;
                    }

                    if (library == null)
                    {
                        // 2️⃣ CREATE NEW LIBRARY
                        // Strip leading slashes and use only last segment as internal URL name
                        string internalName = dto.Url.TrimStart('/').Split('/').Last();

                        var creationInfo = new ListCreationInformation
                        {
                            Title = dto.Name,
                            Url = internalName,
                            TemplateType = (int)ListTemplateType.DocumentLibrary
                        };

                        library = web.Lists.Add(creationInfo);
                        ctx.ExecuteQuery();

                        // Enable versioning
                        library.EnableVersioning = true;
                        library.MajorVersionLimit = 100;
                        library.Update();
                        ctx.ExecuteQuery();
                    }

                    // 3️⃣ ENABLE CONTENT TYPES (required for custom columns)
                    library.ContentTypesEnabled = true;
                    library.Update();
                    ctx.ExecuteQuery();

                    // 4️⃣ ADD / ENSURE METADATA FIELDS
                    if (dto.Fields != null)
                    {
                        // Load existing field internal names to avoid duplicates
                        ctx.Load(library.Fields, flds => flds.Include(f => f.InternalName));
                        ctx.ExecuteQuery();

                        var existingNames = new HashSet<string>(
                            library.Fields.Select(f => f.InternalName),
                            StringComparer.OrdinalIgnoreCase
                        );

                        foreach (var field in dto.Fields)
                        {
                            if (string.IsNullOrWhiteSpace(field.Name)) continue;

                            string internalFieldName = field.Name.Replace(" ", "");
                            if (existingNames.Contains(internalFieldName)) continue; // already exists

                            try
                            {
                                string fieldXml = BuildFieldXml(field);
                                library.Fields.AddFieldAsXml(fieldXml, true, AddFieldOptions.DefaultValue);
                                library.Update();
                                ctx.ExecuteQuery();

                                existingNames.Add(internalFieldName);
                            }
                            catch
                            {
                                // Field may conflict — skip gracefully
                            }
                        }
                    }

                    return Json(new { success = true });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Removes a document library from SharePoint by its server-relative URL.
        /// Called from the Library Manager view when the user confirms deletion.
        /// </summary>
        [HttpPost]
        public ActionResult RemoveLibrary(DeleteLibraryDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.Url))
                return Json(new { success = false, message = "Library URL is required." });

            try
            {
                using (var ctx = helper.GetContext())
                {
                    var web = ctx.Web;

                    // Load all document libraries so we can match by server-relative URL
                    var lists = web.Lists;
                    ctx.Load(lists, ls => ls.Include(
                        l => l.Title,
                        l => l.BaseTemplate,
                        l => l.RootFolder.ServerRelativeUrl
                    ));
                    ctx.ExecuteQuery();

                    var library = lists.FirstOrDefault(l =>
                        l.BaseTemplate == 101 &&
                        l.RootFolder.ServerRelativeUrl.Equals(dto.Url, StringComparison.OrdinalIgnoreCase)
                    );

                    if (library == null)
                        return Json(new { success = false, message = "Library not found in SharePoint." });

                    library.DeleteObject();
                    ctx.ExecuteQuery();

                    return Json(new { success = true });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // ─── Field XML builder ───────────────────────────────────────────────────
        private string BuildFieldXml(MetaFieldDto field)
        {
            string spType;
            switch (field.Type?.ToLower())
            {
                case "number": spType = "Number"; break;
                case "date": spType = "DateTime"; break;
                case "choice": spType = "Choice"; break;
                case "boolean": spType = "Boolean"; break;
                default: spType = "Text"; break;
            }

            string required = field.Required ? "TRUE" : "FALSE";
            string internalName = field.Name.Replace(" ", "");
            string displayName = System.Security.SecurityElement.Escape(field.Name);

            if (spType == "Choice")
            {
                return $@"<Field Type='Choice' DisplayName='{displayName}' Name='{internalName}' Required='{required}'>
    <CHOICES>
        <CHOICE>Option 1</CHOICE>
        <CHOICE>Option 2</CHOICE>
    </CHOICES>
</Field>";
            }

            return $"<Field Type='{spType}' DisplayName='{displayName}' Name='{internalName}' Required='{required}' />";
        }
    }

    // ═══════════════════════════════════════════════
    //  DTOs  (kept in same file for simplicity)
    // ═══════════════════════════════════════════════

    /// <summary>Payload sent from the Library Manager view when saving a library.</summary>
    public class LibraryDto
    {
        public string Name { get; set; }
        public string Url { get; set; }
        public string Icon { get; set; }
        public string Access { get; set; }
        public string Desc { get; set; }
        public List<MetaFieldDto> Fields { get; set; }
    }

    /// <summary>A single metadata field definition.</summary>
    public class MetaFieldDto
    {
        /// <summary>Display name / label for the field.</summary>
        public string Name { get; set; }

        /// <summary>One of: text | number | date | choice | boolean</summary>
        public string Type { get; set; }

        /// <summary>Whether the field is mandatory on upload.</summary>
        public bool Required { get; set; }
    }

    /// <summary>Payload for the RemoveLibrary action.</summary>
    public class DeleteLibraryDto
    {
        /// <summary>Server-relative URL of the library to delete (e.g. /sites/company/HR).</summary>
        public string Url { get; set; }
    }
}
