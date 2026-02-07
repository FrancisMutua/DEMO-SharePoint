using System;
using System.Collections.Generic;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using KCAU_SharePoint.Models;
using Microsoft.SharePoint.Client;

namespace KCAU_SharePoint.Controllers
{
    public class HomeController : Controller
    {
        private string siteUrl = "http://41.89.240.139/sites/edms";
        private string username = "Sharepoint";
        private string password = "Directory@2024";
        private string domain = "UOEMDOMAIN";

        public ActionResult Index()
        {
            var libraries = GetDocumentLibraries();
            ViewBag.Libraries = libraries;
            var model = new DashboardViewModel
            {
                RecentActivities = new List<RecentActivity>()
            };

            using (var ctx = new ClientContext(siteUrl))
            {
                ctx.Credentials = new NetworkCredential(
                          username, GetSecureString(password), domain
                      );
                Web web = ctx.Web;
                ListCollection lists = web.Lists;
                ctx.Load(lists, l => l.Include(
                    list => list.Title,
                    list => list.BaseTemplate,
                    list => list.ItemCount,
                    list => list.Hidden
                ));
                ctx.ExecuteQuery();

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
                    ctx.Load(items, i => i.Include(
                        it => it["Author"],
                        it => it["Editor"],
                        it => it["FileLeafRef"],
                        it => it["Created"]
                    ));
                    ctx.ExecuteQuery();

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
            var libraries = GetDocumentLibraries();
            ViewBag.Libraries = libraries;

            if (string.IsNullOrEmpty(libraryUrl))
                return RedirectToAction("Index");

            ViewBag.LibraryUrl = libraryUrl;

            // Get all files and folders
            var items = GetFolderContents(libraryUrl);
            ViewBag.Items = items;

            return View();
        }
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

                using (var ctx = new ClientContext(siteUrl))
                {
                    ctx.Credentials = new NetworkCredential(
                        username, GetSecureString(password), domain
                    );

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
        public JsonResult RenameItem(string itemUrl, bool isFolder, string newName)
        {
            try
            {
                using (var ctx = new ClientContext(siteUrl))
                {
                    ctx.Credentials = new NetworkCredential(
                        username, GetSecureString(password), domain
                    );

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

                using (var ctx = new ClientContext(siteUrl))
                {
                    ctx.Credentials = new NetworkCredential(username, GetSecureString(password), domain);

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
        public JsonResult UploadFile(HttpPostedFileBase file, string libraryUrl)
        {
            try
            {
                if (file == null || file.ContentLength == 0)
                    return Json(new { success = false, message = "No file received" });

                if (string.IsNullOrWhiteSpace(libraryUrl))
                    return Json(new { success = false, message = "Library URL missing" });

                using (var ctx = new ClientContext(siteUrl))
                {
                    ctx.Credentials = new NetworkCredential(
                        username, GetSecureString(password), domain
                    );

                    var folder = ctx.Web.GetFolderByServerRelativeUrl(libraryUrl);

                    ctx.Load(folder);
                    ctx.ExecuteQuery();

                    var fileInfo = new FileCreationInformation
                    {
                        ContentStream = file.InputStream,
                        Url = file.FileName,
                        Overwrite = true
                    };

                    var upload = folder.Files.Add(fileInfo);
                    ctx.Load(upload);
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

        private List<SPLibrary> GetDocumentLibraries()
        {
            var libraries = new List<SPLibrary>();

            using (var ctx = new ClientContext(siteUrl))
            {
                ctx.Credentials = new System.Net.NetworkCredential(
                    username, GetSecureString(password), domain
                );

                Web web = ctx.Web;
                ListCollection lists = web.Lists;

                ctx.Load(lists, l => l.Include(
                    list => list.Title,
                    list => list.RootFolder.ServerRelativeUrl,
                    list => list.BaseTemplate,
                    list => list.Hidden
                ));
                ctx.ExecuteQuery();

                foreach (var list in lists)
                {
                    if (list.BaseTemplate == 101 && !list.Hidden)
                    {
                        libraries.Add(new SPLibrary
                        {
                            Title = list.Title,
                            Url = list.RootFolder.ServerRelativeUrl
                        });
                    }
                }
            }

            return libraries;
        }
        private List<SPItem> GetFolderContents(string folderUrl)
        {
            var items = new List<SPItem>();

            using (var ctx = new ClientContext(siteUrl))
            {
                ctx.Credentials = new System.Net.NetworkCredential(
                                   username, GetSecureString(password), domain
                               );

                // Load folder, subfolders, and files
                var folder = ctx.Web.GetFolderByServerRelativeUrl(folderUrl);
                ctx.Load(folder);
                ctx.Load(folder.Folders, f => f.Include(ff => ff.Name, ff => ff.ServerRelativeUrl, ff => ff.ListItemAllFields));
                ctx.Load(folder.Files, f => f.Include(ff => ff.Name, ff => ff.ServerRelativeUrl, ff => ff.Length, ff => ff.ListItemAllFields));
                ctx.ExecuteQuery();

                // Initialize visibleFields to empty
                HashSet<string> visibleFields = new HashSet<string>();

                try
                {
                    // Attempt to find the parent list by folder URL
                    var lists = ctx.Web.Lists;
                    ctx.Load(lists);
                    ctx.ExecuteQuery();

                    List parentList = null;
                    foreach (var list in lists)
                    {
                        // Check if the folder belongs to this list
                        if (folderUrl.StartsWith(list.RootFolder.ServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase))
                        {
                            parentList = list;
                            break;
                        }
                    }

                    if (parentList != null)
                    {
                        // Load the default view
                        var view = parentList.DefaultView;
                        ctx.Load(view, v => v.ViewFields);
                        ctx.ExecuteQuery();

                        visibleFields = new HashSet<string>(view.ViewFields);
                    }
                }
                catch
                {
                    visibleFields = new HashSet<string>();
                }

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

                        if (subFolder.ListItemAllFields != null)
                        {
                            foreach (var field in subFolder.ListItemAllFields.FieldValues)
                            {
                                if (visibleFields.Contains(field.Key))
                                    spItem.Metadata[field.Key] = field.Value?.ToString() ?? "";
                            }
                        }

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

                    if (file.ListItemAllFields != null)
                    {
                        foreach (var field in file.ListItemAllFields.FieldValues)
                        {
                            if (visibleFields.Contains(field.Key))
                                spItem.Metadata[field.Key] = field.Value?.ToString() ?? "";
                        }
                    }

                    items.Add(spItem);
                }
            }

            return items;
        }
        [HttpPost]
        public JsonResult ViewDocument(string fileUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(fileUrl))
                    return Json(new { success = false, message = "Invalid file url" });

                using (var ctx = new ClientContext(siteUrl))
                {
                    ctx.Credentials = new System.Net.NetworkCredential(
                    username, GetSecureString(password), domain
                );
                    var file = ctx.Web.GetFileByServerRelativeUrl(fileUrl);
                    ctx.Load(file);
                    ctx.ExecuteQuery();

                    var stream = file.OpenBinaryStream();
                    ctx.ExecuteQuery();

                    byte[] bytes;
                    using (var ms = new MemoryStream())
                    {
                        stream.Value.CopyTo(ms);
                        bytes = ms.ToArray();
                    }
                    var base1 = Convert.ToBase64String(bytes);
                    var type1 = MimeMapping.GetMimeMapping(file.Name);
                    return Json(new
                    {
                        success = true,
                        base64 = Convert.ToBase64String(bytes),
                        contentType = "application/pdf" // correct MIME type for PDF
                    });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }
        public JsonResult GetFileVersions(string fileUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(fileUrl))
                    return Json(new { success = false, message = "Invalid file URL" });

                using (var ctx = new ClientContext(siteUrl))
                {
                    ctx.Credentials = new System.Net.NetworkCredential(
                     username, GetSecureString(password), domain
                 );
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


        private System.Security.SecureString GetSecureString(string str)
        {
            var secure = new System.Security.SecureString();
            foreach (char c in str)
                secure.AppendChar(c);
            return secure;
        }

    }


}