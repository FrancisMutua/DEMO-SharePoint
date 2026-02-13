using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using KCAU_SharePoint.Models;
using Microsoft.SharePoint.ApplicationPages.ClientPickerQuery;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Utilities;

namespace KCAU_SharePoint.Controllers
{
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
                var instance = helper.CreateWorkflowInstance(itemUrl, itemName, User.Identity.Name);

                return Json(new { success = true, workflowInstanceId = instance.Id });
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
        public JsonResult UploadFile(HttpPostedFileBase file, string libraryUrl)
        {
            try
            {
                if (file == null || file.ContentLength == 0)
                    return Json(new { success = false, message = "No file received" });

                if (string.IsNullOrWhiteSpace(libraryUrl))
                    return Json(new { success = false, message = "Library URL missing" });

                using (var ctx = helper.GetContext())
                {
                    
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

        
        private List<SPItem> GetFolderContents(string folderUrl)
        {
            var items = new List<SPItem>();

            using (var ctx = helper.GetContext())
            {

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

                using (var ctx = helper.GetContext())
                {
                   
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
        private void SendShareNotification(ClientContext ctx,string toEmail,string displayName,string itemLink,string role)
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

            // IMPORTANT: escaped for C# string
            return $"{loginName}";
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

    }


}