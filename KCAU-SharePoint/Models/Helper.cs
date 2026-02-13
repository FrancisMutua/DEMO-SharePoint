using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Net;
using System.Web;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using Microsoft.SharePoint.News.DataModel;
using Newtonsoft.Json;

namespace KCAU_SharePoint.Models
{

    public static class ClientContextExtensions
    {
        public static List GetListByUrl(this Web web, string itemUrl)
        {
            // Extract list relative URL from item URL
            var siteRelativeUrl = web.ServerRelativeUrl.TrimEnd('/');
            var listRelativeUrl = itemUrl.Substring(itemUrl.IndexOf(siteRelativeUrl) + siteRelativeUrl.Length);
            var listUrlParts = listRelativeUrl.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var listRootUrl = siteRelativeUrl + "/" + string.Join("/", listUrlParts.Take(2)); // /site/Lists/ListName or /site/DocumentLibrary

            return web.GetList(listRootUrl);
        }
    }

    public class Helper
    {
        private readonly string siteUrl;
        private readonly string username;
        private readonly string password;
        private readonly string domain;
        // Use the same static storage as your controller/demo
        private static List<WorkflowModel> _workflows = new List<WorkflowModel>();
        private static List<WorkflowInstance> _workflowInstances = new List<WorkflowInstance>();
        private static List<SPItem> _items = new List<SPItem>();

        public Helper()
        {
            siteUrl = ConfigurationManager.AppSettings["siteUrl"];
            username = ConfigurationManager.AppSettings["username"];
            password = ConfigurationManager.AppSettings["password"];
            domain = ConfigurationManager.AppSettings["domain"];
        }

        public ClientContext GetContext(string username = null, string password = null)
        {
            string user = username ?? HttpContext.Current.Session["Username"]?.ToString();
            string pass = password ?? HttpContext.Current.Session["password"]?.ToString();

            if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
                throw new InvalidOperationException("User is not logged in.");

            var ctx = new ClientContext(siteUrl);
            ctx.Credentials = new NetworkCredential(user, pass, domain);
            ctx.ExecutingWebRequest += (sender, e) =>
            {
                e.WebRequestExecutor.WebRequest.PreAuthenticate = true;
            };
            return ctx;
        }
        public void CreateWorkflow(CreateWorkflowRequest model)
        {
            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("WorkflowConfigs");

                var itemCreateInfo = new ListItemCreationInformation();
                var item = list.AddItem(itemCreateInfo);

                item["Title"] = model.Name;
                item["LibraryUrl"] = model.LibraryUrl;
                item["Levels"] = model.Levels;
                item["IsActive"] = true;

                item["StagesJson"] = JsonConvert.SerializeObject(model.Stages);

                item.Update();
                ctx.ExecuteQuery();
            }
        }
        public void DeleteWorkflow(int id)
        {
            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("WorkflowConfigs");
                var item = list.GetItemById(id);

                item.DeleteObject();
                ctx.ExecuteQuery();
            }
        }

        public PrincipalContext GetPrincipalContext()
        {
            return new PrincipalContext(ContextType.Domain, domain, username, password);
        }

        // Get approvers for a specific workflow level
        public List<string> GetApprovers(int workflowId, int level)
        {
            var emails = new List<string>();

            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("WorkflowLevels");

                var caml = $@"
        <View>
            <Query>
                <Where>
                    <And>
                        <Eq>
                            <FieldRef Name='WorkflowId'/>
                            <Value Type='Number'>{workflowId}</Value>
                        </Eq>
                        <Eq>
                            <FieldRef Name='LevelNo'/>
                            <Value Type='Number'>{level}</Value>
                        </Eq>
                    </And>
                </Where>
            </Query>
        </View>";

                var items = list.GetItems(new CamlQuery { ViewXml = caml });

                ctx.Load(items, i => i.Include(item => item["Approvers"]));
                ctx.ExecuteQuery();

                var userIds = new HashSet<int>();

                foreach (var item in items)
                {
                    var people = item["Approvers"] as FieldUserValue[];
                    if (people == null) continue;

                    foreach (var p in people)
                        userIds.Add(p.LookupId);
                }

                if (userIds.Any())
                {
                    foreach (var id in userIds)
                    {
                        var user = ctx.Web.GetUserById(id);
                        ctx.Load(user);
                    }

                    ctx.ExecuteQuery(); // Single batch call

                    foreach (var id in userIds)
                    {
                        var user = ctx.Web.GetUserById(id);
                        if (!string.IsNullOrEmpty(user.Email))
                            emails.Add(user.Email);
                    }
                }
            }

            return emails;
        }

        public WorkflowModel GetWorkflowForLibrary(string libraryUrl)
        {
            if (string.IsNullOrEmpty(libraryUrl))
                return null;

            // Assume only one workflow per library for simplicity
            return _workflows.FirstOrDefault(w => w.LibraryUrl.Equals(libraryUrl, StringComparison.OrdinalIgnoreCase));
        }
        public string GetLibraryUrlFromItem(string itemUrl)
        {
            if (string.IsNullOrEmpty(itemUrl))
                throw new ArgumentException("Item URL cannot be empty");

            // Example: "/Documents/Finance/Doc1.pdf" -> "/Documents"
            var parts = itemUrl.Trim('/').Split('/');
            if (parts.Length == 0)
                throw new Exception("Invalid item URL format");

            return "/" + parts[0];
        }

        // Create a workflow instance when a document is submitted for approval
        public WorkflowInstance CreateWorkflowInstance(string itemUrl, string itemName, string submittedBy)
        {
            // Dynamically get library from the item URL
            string libraryUrl = GetLibraryUrlFromItem(itemUrl);

            var workflow = GetWorkflowForLibrary(libraryUrl);
            if (workflow == null)
                throw new Exception("No workflow configured for this library.");

            var instance = new WorkflowInstance
            {
                Id = _workflowInstances.Count + 1,
                ItemUrl = itemUrl,
                ItemName = itemName,
                WorkflowId = workflow.Id,
                CurrentLevel = 1,
                Status = "Pending",
                SubmittedBy = submittedBy,
                SubmittedDate = DateTime.Now,
                TotalLevels = workflow.Levels
            };

            _workflowInstances.Add(instance);
            return instance;
        }

        // Optional: Helper to get instances for a user
        public List<WorkflowInstance> GetPendingApprovalsForUser(string userName)
        {
            var pending = new List<WorkflowInstance>();

            foreach (var instance in _workflowInstances.Where(i => i.Status == "Pending"))
            {
                var workflow = _workflows.FirstOrDefault(w => w.Id == instance.WorkflowId);
                var stage = workflow?.Stages.FirstOrDefault(s => s.Level == instance.CurrentLevel);

                if (stage != null && stage.Approvers.Contains(userName, StringComparer.OrdinalIgnoreCase))
                    pending.Add(instance);
            }

            return pending;
        }

        // Get all workflows
        public List<WorkflowConfigVM> GetWorkflows()
        {
            var result = new List<WorkflowConfigVM>();

            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("WorkflowConfigs");
                var query = CamlQuery.CreateAllItemsQuery();
                var items = list.GetItems(query);

                ctx.Load(items);
                ctx.ExecuteQuery();

                foreach (var item in items)
                {
                    var stagesJson = item["StagesJson"]?.ToString();

                    var stages = string.IsNullOrEmpty(stagesJson)
                        ? new List<WorkflowStageModel>()
                        : JsonConvert.DeserializeObject<List<WorkflowStageModel>>(stagesJson);

                    result.Add(new WorkflowConfigVM
                    {
                        Id = item.Id,
                        WorkflowName = item["Title"]?.ToString(),
                        LibraryUrl = item["LibraryUrl"]?.ToString(),
                        Levels = Convert.ToInt32(item["Levels"]),
                        IsActive = (bool)item["IsActive"],
                        Stages = stages   // ✅ IMPORTANT
                    });
                }
            }

            return result;
        }


        // Get all document libraries
        public List<SPLibrary> GetDocumentLibraries()
        {
            var libraries = new List<SPLibrary>();

            using (var ctx = GetContext())
            {
                var lists = ctx.Web.Lists;

                ctx.Load(lists, l => l.Include(
                    list => list.Title,
                    list => list.RootFolder.ServerRelativeUrl,
                    list => list.BaseTemplate,
                    list => list.Hidden,
                    list => list.EffectiveBasePermissions
                ));

                ctx.ExecuteQuery();

                foreach (var list in lists)
                {
                    if (list.BaseTemplate == 101 &&
                        !list.Hidden &&
                        list.EffectiveBasePermissions.Has(PermissionKind.ViewListItems))
                    {
                        libraries.Add(new SPLibrary
                        {
                            Title = list.Title,
                            Url = list.RootFolder.ServerRelativeUrl
                        });
                    }
                }
            }

            return libraries.OrderBy(l => l.Title).ToList();
        }


        // Convert string to SecureString
        public System.Security.SecureString GetSecureString(string str)
        {
            var secure = new System.Security.SecureString();
            foreach (char c in str)
                secure.AppendChar(c);
            return secure;
        }
    }

}