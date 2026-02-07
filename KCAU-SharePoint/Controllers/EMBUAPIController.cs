using System;
using System.Linq;
using System.Net;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;

namespace KCAU_SharePoint.Controllers
{
    [RoutePrefix("api/sharepoint")]
    public class EMBUAPIController : Controller
    {
        private const string SharePointSiteUrl = "http://41.89.240.139/sites/edms";
        private const string DocumentLibraryName = "Applications Attachment";
        private const string SharePointUsername = "Sharepoint";
        private const string SharePointPassword = "Directory@2024";
        private const string SharePointDomain = "UOEMDOMAIN";

        [HttpPost]
        [Route("upload")]
        public ActionResult UploadFile(UploadRequest model)
        {
            if (model == null ||
                string.IsNullOrEmpty(model.FileName) ||
                string.IsNullOrEmpty(model.Base64Content))
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest, "Invalid request");
            }

            try
            {
                using (var context = new ClientContext(SharePointSiteUrl))
                {
                    context.Credentials = new NetworkCredential(
                        SharePointUsername,
                        SharePointPassword,
                        SharePointDomain
                    );

                    var list = context.Web.Lists.GetByTitle(DocumentLibraryName);
                    context.Load(list.RootFolder);
                    context.ExecuteQuery();

                    // Ensure folder exists
                    var folder = EnsureFolder(context, list.RootFolder, model.FolderName);

                    // Block specific file
                    if (model.FileName.Equals(
                        "Emergency Medical Procedure Preauthorization Form",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return Json(new { success = true, message = "File skipped (blocked)" });
                    }

                    // Check if file exists
                    context.Load(folder.Files);
                    context.ExecuteQuery();

                    if (folder.Files.Any(f =>
                        f.Name.Equals(model.FileName, StringComparison.OrdinalIgnoreCase)))
                    {
                        return Json(new { success = true, message = "File already exists, skipped" });
                    }

                    // Upload file
                    byte[] fileBytes = Convert.FromBase64String(model.Base64Content);

                    var fileInfo = new FileCreationInformation
                    {
                        Content = fileBytes,
                        Url = model.FileName,
                        Overwrite = false
                    };

                    folder.Files.Add(fileInfo);
                    context.ExecuteQuery();

                    return Json(new { success = true, message = "File uploaded successfully" });
                }
            }
            catch (Exception ex)
            {
                return new HttpStatusCodeResult(
                    HttpStatusCode.InternalServerError,
                    ex.Message
                );
            }
        }

        private Folder EnsureFolder(ClientContext context, Folder rootFolder, string folderName)
        {
            folderName = folderName.Replace("/", "-");

            context.Load(rootFolder.Folders);
            context.ExecuteQuery();

            var existingFolder = rootFolder.Folders
                .FirstOrDefault(f =>
                    f.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase));

            if (existingFolder != null)
                return existingFolder;

            var newFolder = rootFolder.Folders.Add(folderName);
            context.Load(newFolder);
            context.ExecuteQuery();

            return newFolder;
        }
    }

    public class UploadRequest
    {
        public string FolderName { get; set; }
        public string FileName { get; set; }
        public string Base64Content { get; set; }
    }
}
