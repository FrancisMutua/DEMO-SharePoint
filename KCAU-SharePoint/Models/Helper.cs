using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Net;
using System.Runtime.Remoting.Messaging;
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

        public Helper()
        {
            siteUrl = ConfigurationManager.AppSettings["siteUrl"];
            username = ConfigurationManager.AppSettings["username"];
            password = ConfigurationManager.AppSettings["password"];
            domain = ConfigurationManager.AppSettings["domain"];
        }
        public string getDomain => domain;

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

        public WorkflowConfigVM GetWorkflowForLibrary(string libraryUrl)
        {
            if (string.IsNullOrWhiteSpace(libraryUrl))
                return null;

            var normalizedLibraryUrl = NormalizeUrl(libraryUrl);
            var workflows = GetWorkflows();
            return workflows.FirstOrDefault(w =>
                NormalizeUrl(w.LibraryUrl) == normalizedLibraryUrl);
        }

        private string NormalizeUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
                return string.Empty;

            return Uri.UnescapeDataString(url)
                      .Trim()
                      .ToLowerInvariant();
        }



        public string GetLibraryUrlFromItem(string itemUrl)
        {
            if (string.IsNullOrWhiteSpace(itemUrl))
                return null;

            using (ClientContext context = GetContext())
            {
                // Ensure proper encoding for SharePoint
                var encodedUrl = Uri.EscapeUriString(itemUrl);

                var file = context.Web.GetFileByServerRelativeUrl(encodedUrl);
                context.Load(file,
                    f => f.ListItemAllFields,
                    f => f.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl);

                context.ExecuteQuery();

                return Uri.UnescapeDataString(
                    file.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl);
            }
        }

        public List<WorkflowInstance> GetWorkflowInstances()
        {
            var instances = new List<WorkflowInstance>();

            using (var context = GetContext())
            {
                var list = context.Web.Lists.GetByTitle("ApprovalInstances");
                var query = CamlQuery.CreateAllItemsQuery();
                var items = list.GetItems(query);

                context.Load(items);
                context.ExecuteQuery();

                foreach (var item in items)
                {
                    var instance = new WorkflowInstance
                    {
                        Id = item.Id.ToString(),
                        ItemName = item["Title"]?.ToString() ?? "",
                        ItemUrl = item["ItemUrl"]?.ToString() ?? "",
                        WorkflowId = item["WorkflowId"]?.ToString() ?? "",
                        CurrentLevel = item["Stage"]?.ToString() ?? "1",
                        Status = item["Status"]?.ToString() ?? "",
                        SubmittedBy = item["SubmittedBy"]?.ToString() ?? "",
                        TotalLevels = item["TotalLevels"]?.ToString() ?? "1",
                        Approver = item["Approver"]?.ToString() ?? ""
                    };

                    
                    instances.Add(instance);
                }
            }

            return instances;
        }


        // Create a workflow instance when a document is submitted for approval
        public void CreateWorkflowInstance(string itemUrl, string itemName, string submittedBy)
        {
            string libraryUrl = GetLibraryUrlFromItem(itemUrl);

            var workflow = GetWorkflowForLibrary(libraryUrl);
            if (workflow == null)
                throw new Exception("No workflow configured for this library.");

            using (ClientContext context = GetContext())
            {
                var list = context.Web.Lists.GetByTitle("ApprovalInstances");

                foreach (var stage in workflow.Stages)
                {
                    foreach (var approver in stage.Approvers)
                    {
                        var itemCreateInfo = new ListItemCreationInformation();
                        var listItem = list.AddItem(itemCreateInfo);

                        listItem["Title"] = itemName;
                        listItem["ItemUrl"] = itemUrl;
                        listItem["WorkflowId"] = workflow.Id;
                        listItem["Stage"] = stage.Level;
                        listItem["Approver"] = approver; // Assign approver for this level
                        listItem["Status"] = stage.Level == 1 ? "Pending" : "Created";
                        listItem["SubmittedBy"] = submittedBy;
                        listItem["TotalLevels"] = workflow.Levels;

                        listItem.Update();
                    }
                }

                context.ExecuteQuery();
            }
        }

        // Helper class to deserialize stages
        public class WorkflowStage
        {
            public int Level { get; set; }
            public List<string> Approvers { get; set; }
        }


        /* // Optional: Helper to get instances for a user
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
         }*/

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