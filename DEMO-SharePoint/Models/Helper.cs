using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Net;
using System.Web;
using DEMO_SharePoint.Models;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;

namespace DEMO_SharePoint.Models
{
    // Extension helpers
    public static class ClientContextExtensions
    {
        public static List GetListByUrl(this Web web, string itemUrl)
        {
            var siteRelativeUrl = web.ServerRelativeUrl.TrimEnd('/');
            var listRelativeUrl = itemUrl.Substring(itemUrl.IndexOf(siteRelativeUrl) + siteRelativeUrl.Length);
            var listUrlParts    = listRelativeUrl.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var listRootUrl     = siteRelativeUrl + "/" + string.Join("/", listUrlParts.Take(2));
            return web.GetList(listRootUrl);
        }
    }

    // -------------------------------------------------------------------
    //  HELPER - SharePoint service layer
    // -------------------------------------------------------------------
    public class Helper
    {
        private readonly string siteUrl;
        private readonly string username;
        private readonly string password;
        private readonly string domain;

        private readonly NotificationService _notify;

        public Helper()
        {
            siteUrl  = ConfigurationManager.AppSettings["siteUrl"];
            username = ConfigurationManager.AppSettings["username"];
            password = ConfigurationManager.AppSettings["password"];
            domain   = ConfigurationManager.AppSettings["domain"];
            _notify  = new NotificationService();
        }

        public string getDomain => domain;

        // Context factory
        /// <summary>
        /// Builds a CSOM ClientContext using the session credentials.
        /// Session keys are "Username" and "Password" (capital U/P) everywhere.
        /// Falls back to the service account from Web.config if session is empty.
        /// </summary>
        public ClientContext GetContext(string overrideUser = null, string overridePass = null)
        {
            string user = overrideUser
                          ?? HttpContext.Current?.Session?["Username"]?.ToString()
                          ?? username;
            string pass = overridePass
                          ?? HttpContext.Current?.Session?["Password"]?.ToString()
                          ?? password;

            if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
                throw new InvalidOperationException("User is not logged in.");

            var ctx = new ClientContext(siteUrl);
            ctx.Credentials = new NetworkCredential(user, pass, domain);
            ctx.ExecutingWebRequest += (s, e) =>
                e.WebRequestExecutor.WebRequest.PreAuthenticate = true;
            return ctx;
        }

        public PrincipalContext GetPrincipalContext() =>
            new PrincipalContext(ContextType.Domain, domain, username, password);

        public System.Security.SecureString GetSecureString(string str)
        {
            var ss = new System.Security.SecureString();
            foreach (char c in str) ss.AppendChar(c);
            return ss;
        }

        // -------------------------------------------------------------------
        //  WORKFLOW CONFIG CRUD
        // -------------------------------------------------------------------

        public void CreateWorkflow(CreateWorkflowRequest model)
        {
            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("WorkflowConfigs");
                var item = list.AddItem(new ListItemCreationInformation());

                item["Title"]             = model.Name;
                item["LibraryUrl"]        = model.LibraryUrl;
                item["Levels"]            = model.Levels;
                item["IsActive"]          = true;
                item["StagesJson"]        = JsonConvert.SerializeObject(model.Stages);
                item["TriggerEvents"]     = JsonConvert.SerializeObject(
                    model.TriggerEvents ?? new List<string> { "Manual" });
                item["RejectionBehavior"] = model.RejectionBehavior ?? "ReturnToSubmitter";
                item["NotifyOnSubmit"]    = model.NotifyOnSubmit;
                item["NotifyOnApprove"]   = model.NotifyOnApprove;
                item["NotifyOnReject"]    = model.NotifyOnReject;
                item["NotifyOnEscalate"]  = model.NotifyOnEscalate;
                item["NotifyOnDelegate"]  = model.NotifyOnDelegate;
                item["NotifyOnComplete"]  = model.NotifyOnComplete;

                item.Update();
                ctx.ExecuteQuery();
            }
        }

        public void DeleteWorkflow(int id)
        {
            using (var ctx = GetContext())
            {
                var item = ctx.Web.Lists.GetByTitle("WorkflowConfigs").GetItemById(id);
                item.DeleteObject();
                ctx.ExecuteQuery();
            }
        }

        public void ToggleWorkflow(int id, bool isActive)
        {
            using (var ctx = GetContext())
            {
                var item = ctx.Web.Lists.GetByTitle("WorkflowConfigs").GetItemById(id);
                ctx.Load(item);
                ctx.ExecuteQuery();
                item["IsActive"] = isActive;
                item.Update();
                ctx.ExecuteQuery();
            }
        }

        public List<WorkflowConfigVM> GetWorkflows()
        {
            var result = new List<WorkflowConfigVM>();
            using (var ctx = GetContext())
            {
                var items = ctx.Web.Lists.GetByTitle("WorkflowConfigs")
                              .GetItems(CamlQuery.CreateAllItemsQuery());
                ctx.Load(items);
                ctx.ExecuteQuery();

                foreach (var item in items)
                {
                    result.Add(new WorkflowConfigVM
                    {
                        Id                = item.Id,
                        WorkflowName      = item["Title"]?.ToString() ?? "",
                        LibraryUrl        = item["LibraryUrl"]?.ToString() ?? "",
                        Levels            = SafeInt(item["Levels"]),
                        IsActive          = SafeBool(item["IsActive"]),
                        Stages            = DeserializeJson<List<WorkflowStageModel>>(
                                               item["StagesJson"]?.ToString())
                                           ?? new List<WorkflowStageModel>(),
                        TriggerEvents     = DeserializeJson<List<string>>(
                                               item["TriggerEvents"]?.ToString())
                                           ?? new List<string> { "Manual" },
                        RejectionBehavior = item["RejectionBehavior"]?.ToString() ?? "ReturnToSubmitter",
                        NotifyOnSubmit    = SafeBool(item["NotifyOnSubmit"]),
                        NotifyOnApprove   = SafeBool(item["NotifyOnApprove"]),
                        NotifyOnReject    = SafeBool(item["NotifyOnReject"]),
                        NotifyOnEscalate  = SafeBool(item["NotifyOnEscalate"]),
                        NotifyOnDelegate  = SafeBool(item["NotifyOnDelegate"]),
                        NotifyOnComplete  = SafeBool(item["NotifyOnComplete"]),
                    });
                }
            }
            return result;
        }

        public WorkflowConfigVM GetWorkflowForLibrary(string libraryUrl)
        {
            if (string.IsNullOrWhiteSpace(libraryUrl)) return null;
            var norm = NormalizeUrl(libraryUrl);
            return GetWorkflows()
                   .FirstOrDefault(w => w.IsActive && NormalizeUrl(w.LibraryUrl) == norm);
        }

        // -------------------------------------------------------------------
        //  WORKFLOW INSTANCE QUERIES
        // -------------------------------------------------------------------

        public List<WorkflowInstance> GetPendingApprovalsForUser(string username)
        {
            return QueryInstances($@"
                <Where>
                  <And>
                    <Eq><FieldRef Name='Approver'/><Value Type='Text'>{username}</Value></Eq>
                    <Eq><FieldRef Name='Status'/><Value Type='Text'>Pending</Value></Eq>
                  </And>
                </Where>");
        }

        public List<WorkflowInstance> GetSubmittedByUser(string username)
        {
            return QueryInstances($@"
                <Where>
                  <Eq><FieldRef Name='SubmittedBy'/><Value Type='Text'>{username}</Value></Eq>
                </Where>",
                orderByField: "SubmittedDate", ascending: false);
        }

        public List<WorkflowInstance> GetInstancesByRunId(string workflowRunId)
        {
            return QueryInstances($@"
                <Where>
                  <Eq><FieldRef Name='WorkflowRunId'/><Value Type='Text'>{workflowRunId}</Value></Eq>
                </Where>");
        }

        public List<WorkflowInstance> GetWorkflowInstances() => QueryInstances(null);

        private List<WorkflowInstance> QueryInstances(string whereClause,
            string orderByField = null, bool ascending = true)
        {
            var result = new List<WorkflowInstance>();
            using (var ctx = GetContext())
            {
                string orderBy = orderByField != null
                    ? $"<OrderBy><FieldRef Name='{orderByField}' Ascending='{(ascending ? "TRUE" : "FALSE")}'/></OrderBy>"
                    : "";
                var caml = new CamlQuery
                {
                    ViewXml = $"<View><Query>{whereClause ?? ""}{orderBy}</Query><RowLimit>500</RowLimit></View>"
                };
                var items = ctx.Web.Lists.GetByTitle("ApprovalInstances").GetItems(caml);
                ctx.Load(items);
                ctx.ExecuteQuery();
                foreach (var item in items)
                    result.Add(MapInstance(item));
            }
            return result;
        }

        private WorkflowInstance MapInstance(ListItem item)
        {
            return new WorkflowInstance
            {
                Id               = item.Id.ToString(),
                WorkflowRunId    = item["WorkflowRunId"]?.ToString() ?? "",
                ItemUrl          = item["ItemUrl"]?.ToString() ?? "",
                ItemName         = item["Title"]?.ToString() ?? "",
                WorkflowId       = item["WorkflowId"]?.ToString() ?? "",
                WorkflowName     = item["WorkflowName"]?.ToString() ?? "",
                CurrentLevel     = item["Stage"]?.ToString() ?? "1",
                TotalLevels      = item["TotalLevels"]?.ToString() ?? "1",
                ApprovalMode     = item["ApprovalMode"]?.ToString() ?? "Any",
                Status           = item["Status"]?.ToString() ?? "",
                Approver         = item["Approver"]?.ToString() ?? "",
                OriginalApprover = item["OriginalApprover"]?.ToString() ?? "",
                SubmittedBy      = item["SubmittedBy"]?.ToString() ?? "",
                SubmittedDate    = SafeDate(item["SubmittedDate"]) ?? DateTime.Now,
                ActionDate       = SafeDate(item["ActionDate"]),
                DueDate          = SafeDate(item["DueDate"]),
                Comments         = item["Comments"]?.ToString() ?? "",
                IsEscalated      = SafeBool(item["IsEscalated"]),
            };
        }

        // -------------------------------------------------------------------
        //  TRIGGER SYSTEM
        // -------------------------------------------------------------------

        /// <summary>
        /// Fires when a document is uploaded or updated.
        /// Returns the WorkflowRunId if a workflow was triggered, null otherwise.
        /// </summary>
        public string ProcessTrigger(string eventType, string itemUrl,
                                      string itemName, string triggeredBy)
        {
            try
            {
                string libraryUrl = GetLibraryUrlFromItem(itemUrl);
                var workflow = GetWorkflowForLibrary(libraryUrl);
                if (workflow == null) return null;
                if (workflow.TriggerEvents == null ||
                    !workflow.TriggerEvents.Any(t =>
                        string.Equals(t, eventType, StringComparison.OrdinalIgnoreCase)))
                    return null;
                return CreateWorkflowRun(itemUrl, itemName, triggeredBy, eventType);
            }
            catch { return null; }
        }

        // -------------------------------------------------------------------
        //  WORKFLOW RUN - CREATE
        // -------------------------------------------------------------------

        /// <summary>
        /// Returns true if an active (Pending or Waiting) workflow run exists for the given item URL.
        /// </summary>
        public bool HasActiveWorkflowRun(string itemUrl)
        {
            using (var ctx = GetContext())
            {
                var caml = new CamlQuery
                {
                    ViewXml = $@"<View><Query><Where>
                        <And>
                          <Eq><FieldRef Name='ItemUrl'/><Value Type='Text'>{itemUrl.Replace("'", "''")}</Value></Eq>
                          <In>
                            <FieldRef Name='Status'/>
                            <Values>
                              <Value Type='Text'>Pending</Value>
                              <Value Type='Text'>Waiting</Value>
                            </Values>
                          </In>
                        </And>
                    </Where><RowLimit>1</RowLimit></Query></View>"
                };
                var list  = ctx.Web.Lists.GetByTitle("ApprovalInstances");
                var items = list.GetItems(caml);
                ctx.Load(items);
                ctx.ExecuteQuery();
                return items.Count > 0;
            }
        }

        /// <summary>
        /// Returns a dictionary of itemUrl -> human-readable approval status
        /// for all items in the given library, derived from the ApprovalInstances list.
        /// Documents are never modified; status is read-only from ApprovalInstances.
        /// </summary>
        public Dictionary<string, string> GetApprovalStatusForLibrary(string libraryUrl)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var ctx = GetContext())
                {
                    var list = ctx.Web.Lists.GetByTitle("ApprovalInstances");
                    var caml = new CamlQuery
                    {
                        ViewXml = $@"<View><Query><Where>
                            <BeginsWith>
                                <FieldRef Name='ItemUrl'/>
                                <Value Type='Text'>{libraryUrl.Replace("'", "''")}</Value>
                            </BeginsWith>
                        </Where></Query><RowLimit>5000</RowLimit></View>"
                    };
                    var spItems = list.GetItems(caml);
                    ctx.Load(spItems);
                    ctx.ExecuteQuery();

                    var rows = new List<(string itemUrl, string runId, int stage,
                                         int total, string status, DateTime submitted)>();
                    foreach (ListItem li in spItems)
                    {
                        rows.Add((
                            itemUrl:   li["ItemUrl"]?.ToString() ?? "",
                            runId:     li["WorkflowRunId"]?.ToString() ?? "",
                            stage:     SafeInt(li["Stage"]),
                            total:     SafeInt(li["TotalLevels"]),
                            status:    li["Status"]?.ToString() ?? "",
                            submitted: li["SubmittedDate"] is DateTime dt ? dt : DateTime.MinValue
                        ));
                    }

                    var byItem = rows
                        .Where(r => !string.IsNullOrEmpty(r.itemUrl))
                        .GroupBy(r => r.itemUrl, StringComparer.OrdinalIgnoreCase);

                    foreach (var itemGroup in byItem)
                    {
                        var byRun = itemGroup.GroupBy(r => r.runId).ToList();
                        // Prefer the active (Pending/Waiting) run; fall back to most recent
                        var chosen = byRun.FirstOrDefault(g =>
                                         g.Any(r => r.status == "Pending" || r.status == "Waiting"))
                                  ?? byRun.OrderByDescending(g => g.Max(r => r.submitted))
                                          .FirstOrDefault();
                        if (chosen == null) continue;
                        result[itemGroup.Key] = DeriveRunStatus(chosen.ToList());
                    }
                }
            }
            catch { }
            return result;
        }

        private string DeriveRunStatus(
            List<(string itemUrl, string runId, int stage, int total, string status, DateTime submitted)> rows)
        {
            if (rows.Any(r => r.status == "Rejected")) return "Rejected";
            if (rows.All(r => r.status == "Cancelled")) return "Recalled";
            if (rows.All(r => r.status == "Approved" || r.status == "Superseded" || r.status == "Cancelled"))
                return "Approved";

            var pending = rows.Where(r => r.status == "Pending").ToList();
            if (pending.Any())
            {
                int cur = pending.Max(r => r.stage);
                int tot = pending.First().total;
                return cur == 1 && tot == 1 ? "Pending" : $"In Progress - Level {cur} of {tot}";
            }
            return "Pending";
        }

        /// <summary>
        /// Creates a new workflow run.
        /// Stage 1 -> Status = "Pending" | Stage 2+ -> Status = "Waiting".
        /// Returns the WorkflowRunId (GUID).
        /// </summary>
        public string CreateWorkflowRun(string itemUrl, string itemName,
                                         string submittedBy, string trigger = "Manual")
        {
            string libraryUrl = GetLibraryUrlFromItem(itemUrl);
            var workflow = GetWorkflowForLibrary(libraryUrl);
            if (workflow == null)
                throw new Exception("No active workflow is configured for this document library.");

            if (HasActiveWorkflowRun(itemUrl))
                throw new InvalidOperationException(
                    "This document already has an active approval in progress. " +
                    "Please wait for it to complete or recall it before resubmitting.");

            string runId = Guid.NewGuid().ToString();

            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("ApprovalInstances");
                foreach (var stage in workflow.Stages.OrderBy(s => s.Level))
                {
                    bool isFirst = stage.Level == 1;
                    string status = isFirst ? "Pending" : "Waiting";
                    DateTime? due = isFirst && stage.DueInDays > 0
                        ? (DateTime?)DateTime.Now.AddDays(stage.DueInDays)
                        : null;

                    foreach (var approver in stage.Approvers)
                    {
                        var li = list.AddItem(new ListItemCreationInformation());
                        li["Title"]            = itemName;
                        li["WorkflowRunId"]    = runId;
                        li["ItemUrl"]          = itemUrl;
                        li["WorkflowId"]       = workflow.Id.ToString();
                        li["WorkflowName"]     = workflow.WorkflowName;
                        li["Stage"]            = stage.Level;
                        li["TotalLevels"]      = workflow.Levels;
                        li["ApprovalMode"]     = stage.ApprovalMode ?? "Any";
                        li["Approver"]         = approver;
                        li["OriginalApprover"] = "";
                        li["Status"]           = status;
                        li["SubmittedBy"]      = submittedBy;
                        li["SubmittedDate"]    = DateTime.Now;
                        li["DueDate"]          = due.HasValue ? (object)due.Value : null;
                        li["IsEscalated"]      = false;
                        li["Comments"]         = "";
                        li.Update();
                    }
                }
                ctx.ExecuteQuery();
            }

            LogAuditEntry(runId, itemUrl, itemName, workflow.WorkflowName,
                          submittedBy, "Submitted", 1, $"Submitted via {trigger}");

            if (workflow.NotifyOnSubmit)
            {
                string baseUrl = GetSiteBaseUrl();
                var firstStage = workflow.Stages.OrderBy(s => s.Level).FirstOrDefault();
                if (firstStage != null)
                    foreach (var ap in firstStage.Approvers)
                        _notify.NotifySubmitted(ResolveEmail(ap), itemName,
                                                submittedBy, 1, baseUrl, itemUrl);
            }

            return runId;
        }

        // -------------------------------------------------------------------
        //  WORKFLOW RUN - APPROVE
        // -------------------------------------------------------------------

        public void ApproveInstance(string instanceId, string approverUsername, string comments)
        {
            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("ApprovalInstances");
                var item = list.GetItemById(int.Parse(instanceId));
                ctx.Load(item);
                ctx.ExecuteQuery();

                string currentApprover = item["Approver"]?.ToString() ?? "";
                if (!currentApprover.Equals(approverUsername, StringComparison.OrdinalIgnoreCase))
                    throw new UnauthorizedAccessException("You are not the assigned approver for this instance.");

                string status = item["Status"]?.ToString();
                if (status != "Pending")
                    throw new InvalidOperationException($"This instance cannot be approved (status: {status}).");

                int stage       = SafeInt(item["Stage"]);
                int totalLevels = SafeInt(item["TotalLevels"]);
                string runId    = item["WorkflowRunId"]?.ToString();
                string mode     = item["ApprovalMode"]?.ToString() ?? "Any";
                string itemName = item["Title"]?.ToString() ?? "";
                string itemUrl  = item["ItemUrl"]?.ToString() ?? "";
                string wfName   = item["WorkflowName"]?.ToString() ?? "";
                string wfId     = item["WorkflowId"]?.ToString() ?? "";
                string submitter= item["SubmittedBy"]?.ToString() ?? "";

                item["Status"]     = "Approved";
                item["ActionDate"] = DateTime.Now;
                item["Comments"]   = comments ?? "";
                item.Update();
                ctx.ExecuteQuery();

                LogAuditEntry(runId, itemUrl, itemName, wfName, approverUsername,
                              "Approved", stage, comments);

                bool stageComplete = IsStageComplete(ctx, list, runId, stage, mode);

                if (stageComplete)
                {
                    if (string.Equals(mode, "Any", StringComparison.OrdinalIgnoreCase))
                        SupersedePendingInStage(ctx, list, runId, stage, instanceId);

                    var wf = int.TryParse(wfId, out int wfIdInt)
                             ? GetWorkflows().FirstOrDefault(w => w.Id == wfIdInt)
                             : null;

                    if (stage < totalLevels)
                    {
                        ActivateStage(ctx, list, runId, stage + 1);
                        LogAuditEntry(runId, itemUrl, itemName, wfName, "System",
                                      "StageAdvanced", stage + 1,
                                      $"Advanced from Level {stage} to Level {stage + 1}");

                        if (wf?.NotifyOnApprove == true)
                        {
                            _notify.NotifyApproved(ResolveEmail(submitter), itemName,
                                                   approverUsername, stage, totalLevels);
                            var nextStage = wf.Stages.FirstOrDefault(s => s.Level == stage + 1);
                            if (nextStage != null)
                            {
                                string baseUrl = GetSiteBaseUrl();
                                foreach (var ap in nextStage.Approvers)
                                    _notify.NotifySubmitted(ResolveEmail(ap), itemName,
                                                            submitter, stage + 1, baseUrl, itemUrl);
                            }
                        }
                    }
                    else
                    {
                        LogAuditEntry(runId, itemUrl, itemName, wfName, "System",
                                      "Completed", stage, "All approval levels satisfied.");

                        if (wf?.NotifyOnComplete == true)
                        {
                            _notify.NotifyApproved(ResolveEmail(submitter), itemName,
                                                   approverUsername, stage, totalLevels);
                            _notify.NotifyCompleted(ResolveEmail(submitter), itemName, totalLevels);
                        }
                    }
                }
                else if (string.Equals(mode, "All", StringComparison.OrdinalIgnoreCase))
                {
                    var wf = int.TryParse(wfId, out int wfIdInt)
                             ? GetWorkflows().FirstOrDefault(w => w.Id == wfIdInt)
                             : null;
                    if (wf?.NotifyOnApprove == true)
                        _notify.NotifyApproved(ResolveEmail(submitter), itemName,
                                               approverUsername, stage, totalLevels);
                }
            }
        }

        // -------------------------------------------------------------------
        //  WORKFLOW RUN - REJECT
        // -------------------------------------------------------------------

        public void RejectInstance(string instanceId, string approverUsername, string comments)
        {
            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("ApprovalInstances");
                var item = list.GetItemById(int.Parse(instanceId));
                ctx.Load(item);
                ctx.ExecuteQuery();

                string currentApprover = item["Approver"]?.ToString() ?? "";
                if (!currentApprover.Equals(approverUsername, StringComparison.OrdinalIgnoreCase))
                    throw new UnauthorizedAccessException("You are not the assigned approver.");

                if (item["Status"]?.ToString() != "Pending")
                    throw new InvalidOperationException("Only Pending instances can be rejected.");

                int stage       = SafeInt(item["Stage"]);
                int totalLevels = SafeInt(item["TotalLevels"]);
                string runId    = item["WorkflowRunId"]?.ToString();
                string itemName = item["Title"]?.ToString() ?? "";
                string itemUrl  = item["ItemUrl"]?.ToString() ?? "";
                string wfName   = item["WorkflowName"]?.ToString() ?? "";
                string wfId     = item["WorkflowId"]?.ToString() ?? "";
                string submitter= item["SubmittedBy"]?.ToString() ?? "";

                item["Status"]     = "Rejected";
                item["ActionDate"] = DateTime.Now;
                item["Comments"]   = comments ?? "";
                item.Update();
                ctx.ExecuteQuery();

                LogAuditEntry(runId, itemUrl, itemName, wfName, approverUsername,
                              "Rejected", stage, comments);

                string behavior = "ReturnToSubmitter";
                WorkflowConfigVM wf = null;
                if (int.TryParse(wfId, out int wfIdInt))
                {
                    wf = GetWorkflows().FirstOrDefault(w => w.Id == wfIdInt);
                    if (wf != null) behavior = wf.RejectionBehavior ?? "ReturnToSubmitter";
                }

                switch (behavior)
                {
                    case "ReturnToPreviousLevel":
                        if (stage > 1)
                        {
                            CancelStageInstances(ctx, list, runId, stage);
                            ActivateStage(ctx, list, runId, stage - 1);
                            LogAuditEntry(runId, itemUrl, itemName, wfName, "System",
                                          "StageAdvanced", stage - 1,
                                          $"Returned to Level {stage - 1} after rejection at Level {stage}");
                        }
                        else
                        {
                            CancelAllPending(ctx, list, runId);
                            LogAuditEntry(runId, itemUrl, itemName, wfName, "System",
                                          "Cancelled", stage, "Rejected at Level 1 - cancelled.");
                        }
                        break;

                    case "RestartWorkflow":
                        CancelAllPending(ctx, list, runId);
                        ActivateStage(ctx, list, runId, 1);
                        LogAuditEntry(runId, itemUrl, itemName, wfName, "System",
                                      "Restarted", 1,
                                      $"Restarted from Level 1 after rejection at Level {stage}.");
                        break;

                    default: // ReturnToSubmitter
                        CancelAllPending(ctx, list, runId);
                        LogAuditEntry(runId, itemUrl, itemName, wfName, "System",
                                      "Cancelled", stage, "Cancelled - returned to submitter.");
                        break;
                }

                if (wf?.NotifyOnReject == true)
                    _notify.NotifyRejected(ResolveEmail(submitter), itemName,
                                           approverUsername, stage, comments);
            }
        }

        // -------------------------------------------------------------------
        //  WORKFLOW RUN - DELEGATE
        // -------------------------------------------------------------------

        public void DelegateInstance(string instanceId, string fromUser,
                                      string toUser, string reason)
        {
            using (var ctx = GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("ApprovalInstances");
                var item = list.GetItemById(int.Parse(instanceId));
                ctx.Load(item);
                ctx.ExecuteQuery();

                if (!(item["Approver"]?.ToString() ?? "")
                        .Equals(fromUser, StringComparison.OrdinalIgnoreCase))
                    throw new UnauthorizedAccessException("You are not the assigned approver.");

                if (item["Status"]?.ToString() != "Pending")
                    throw new InvalidOperationException("Only Pending instances can be delegated.");

                int stage       = SafeInt(item["Stage"]);
                string runId    = item["WorkflowRunId"]?.ToString();
                string wfName   = item["WorkflowName"]?.ToString() ?? "";
                string wfId     = item["WorkflowId"]?.ToString() ?? "";
                string itemName = item["Title"]?.ToString() ?? "";
                string itemUrl  = item["ItemUrl"]?.ToString() ?? "";

                if (int.TryParse(wfId, out int wfIdInt))
                {
                    var stageModel = GetWorkflows()
                        .FirstOrDefault(w => w.Id == wfIdInt)
                        ?.Stages?.FirstOrDefault(s => s.Level == stage);
                    if (stageModel != null && !stageModel.AllowDelegation)
                        throw new InvalidOperationException("Delegation is not allowed at this stage.");
                }

                item["OriginalApprover"] = fromUser;
                item["Approver"]         = toUser;
                item["Comments"]         = $"Delegated from {fromUser}: {reason}";
                item.Update();
                ctx.ExecuteQuery();

                LogAuditEntry(runId, itemUrl, itemName, wfName, fromUser,
                              "Delegated", stage, $"Delegated to {toUser}. Reason: {reason}");

                var wf = GetWorkflows().FirstOrDefault(w => w.WorkflowName == wfName);
                if (wf?.NotifyOnDelegate == true)
                    _notify.NotifyDelegated(ResolveEmail(toUser), itemName, fromUser, stage, reason);
            }
        }

        // -------------------------------------------------------------------
        //  WORKFLOW RUN - RECALL
        // -------------------------------------------------------------------

        public void RecallWorkflowRun(string workflowRunId, string requestedBy)
        {
            var instances = GetInstancesByRunId(workflowRunId);
            if (!instances.Any())
                throw new Exception("Workflow run not found.");

            string submitter = instances.First().SubmittedBy;
            if (!submitter.Equals(requestedBy, StringComparison.OrdinalIgnoreCase))
                throw new UnauthorizedAccessException("Only the original submitter may recall this workflow.");

            string itemName = instances.First().ItemName;
            string itemUrl  = instances.First().ItemUrl;
            string wfName   = instances.First().WorkflowName;

            var pendingApprovers = instances
                .Where(i => i.Status == "Pending" || i.Status == "Waiting")
                .Select(i => i.Approver).Distinct().ToList();

            using (var ctx = GetContext())
                CancelAllPending(ctx,
                    ctx.Web.Lists.GetByTitle("ApprovalInstances"),
                    workflowRunId);

            LogAuditEntry(workflowRunId, itemUrl, itemName, wfName, requestedBy,
                          "Recalled", 0, "Recalled by submitter.");

            foreach (var ap in pendingApprovers)
                _notify.NotifyRecalled(ResolveEmail(ap), itemName, requestedBy);
        }

        // -------------------------------------------------------------------
        //  ESCALATION ENGINE
        // -------------------------------------------------------------------

        /// <summary>
        /// Scans all Pending instances past their DueDate and sends escalation
        /// notifications. Returns the count of escalated instances.
        /// </summary>
        public int EscalateOverdueInstances()
        {
            int count = 0;
            using (var ctx = GetContext())
            {
                var caml = new CamlQuery
                {
                    ViewXml = @"<View><Query><Where>
                        <And>
                          <Eq><FieldRef Name='Status'/><Value Type='Text'>Pending</Value></Eq>
                          <Eq><FieldRef Name='IsEscalated'/><Value Type='Boolean'>0</Value></Eq>
                        </And>
                    </Where></Query></View>"
                };

                var list  = ctx.Web.Lists.GetByTitle("ApprovalInstances");
                var items = list.GetItems(caml);
                ctx.Load(items);
                ctx.ExecuteQuery();

                var now = DateTime.Now;
                foreach (var item in items)
                {
                    var due = SafeDate(item["DueDate"]);
                    if (!due.HasValue || due.Value >= now) continue;

                    int stage       = SafeInt(item["Stage"]);
                    string wfId     = item["WorkflowId"]?.ToString() ?? "";
                    string approver = item["Approver"]?.ToString() ?? "";
                    string runId    = item["WorkflowRunId"]?.ToString();
                    string wfName   = item["WorkflowName"]?.ToString() ?? "";
                    string itemName = item["Title"]?.ToString() ?? "";
                    string itemUrl  = item["ItemUrl"]?.ToString() ?? "";
                    string escalateTo = null;

                    if (int.TryParse(wfId, out int wfIdInt))
                    {
                        var wf = GetWorkflows().FirstOrDefault(w => w.Id == wfIdInt);
                        var sm = wf?.Stages?.FirstOrDefault(s => s.Level == stage);
                        escalateTo = sm?.EscalateToEmail;
                        if (wf?.NotifyOnEscalate == true && !string.IsNullOrEmpty(escalateTo))
                            _notify.NotifyEscalated(escalateTo, itemName, approver,
                                                    stage, due.Value);
                    }

                    item["IsEscalated"] = true;
                    item.Update();
                    LogAuditEntry(runId, itemUrl, itemName, wfName, "System",
                                  "Escalated", stage,
                                  $"Overdue - escalated to {escalateTo ?? "N/A"}");
                    count++;
                }
                if (count > 0) ctx.ExecuteQuery();
            }
            return count;
        }

        // -------------------------------------------------------------------
        //  AUDIT LOG
        // -------------------------------------------------------------------

        public void LogAuditEntry(string workflowRunId, string itemUrl,
                                   string itemName, string workflowName,
                                   string actor, string action, int stage, string comments)
        {
            try
            {
                using (var ctx = GetContext())
                {
                    var entry = ctx.Web.Lists.GetByTitle("WorkflowAuditLog")
                                   .AddItem(new ListItemCreationInformation());
                    entry["Title"]         = action;
                    entry["WorkflowRunId"] = workflowRunId;
                    entry["ItemUrl"]       = itemUrl;
                    entry["ItemName"]      = itemName;
                    entry["WorkflowName"]  = workflowName;
                    entry["Actor"]         = actor;
                    entry["Action"]        = action;
                    entry["Stage"]         = stage;
                    entry["Comments"]      = comments ?? "";
                    entry["Timestamp"]     = DateTime.Now;
                    entry.Update();
                    ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("[Audit] {0}", ex.Message);
            }
        }

        public List<WorkflowAuditEntry> GetAuditLog(string itemUrl)
        {
            var result = new List<WorkflowAuditEntry>();
            using (var ctx = GetContext())
            {
                string encoded = itemUrl?.Replace("'", "''") ?? "";
                var caml = new CamlQuery
                {
                    ViewXml = $@"<View><Query>
                        <Where>
                          <Eq><FieldRef Name='ItemUrl'/><Value Type='Text'>{encoded}</Value></Eq>
                        </Where>
                        <OrderBy>
                          <FieldRef Name='Timestamp' Ascending='FALSE'/>
                        </OrderBy>
                    </Query></View>"
                };
                var items = ctx.Web.Lists.GetByTitle("WorkflowAuditLog").GetItems(caml);
                ctx.Load(items);
                ctx.ExecuteQuery();
                foreach (var item in items)
                    result.Add(new WorkflowAuditEntry
                    {
                        Id            = item.Id.ToString(),
                        WorkflowRunId = item["WorkflowRunId"]?.ToString() ?? "",
                        ItemUrl       = item["ItemUrl"]?.ToString() ?? "",
                        ItemName      = item["ItemName"]?.ToString() ?? "",
                        WorkflowName  = item["WorkflowName"]?.ToString() ?? "",
                        Actor         = item["Actor"]?.ToString() ?? "",
                        Action        = item["Action"]?.ToString() ?? "",
                        Stage         = SafeInt(item["Stage"]),
                        Comments      = item["Comments"]?.ToString() ?? "",
                        Timestamp     = SafeDate(item["Timestamp"]) ?? DateTime.Now,
                    });
            }
            return result;
        }

        // -------------------------------------------------------------------
        //  DOCUMENT LIBRARY HELPERS
        // -------------------------------------------------------------------

        public List<SPLibrary> GetDocumentLibraries()
        {
            var libraries = new List<SPLibrary>();
            using (var ctx = GetContext())
            {
                var lists = ctx.Web.Lists;
                ctx.Load(lists, l => l.Include(
                    li => li.Title,
                    li => li.RootFolder.ServerRelativeUrl,
                    li => li.BaseTemplate,
                    li => li.Hidden,
                    li => li.EffectiveBasePermissions));
                ctx.ExecuteQuery();
                foreach (var li in lists)
                {
                    if (li.BaseTemplate == 101 && !li.Hidden &&
                        li.EffectiveBasePermissions.Has(PermissionKind.ViewListItems))
                        libraries.Add(new SPLibrary
                        {
                            Title = li.Title,
                            Url   = li.RootFolder.ServerRelativeUrl
                        });
                }
            }
            return libraries.OrderBy(l => l.Title).ToList();
        }

        public string GetLibraryUrlFromItem(string itemUrl)
        {
            if (string.IsNullOrWhiteSpace(itemUrl)) return null;
            using (var ctx = GetContext())
            {
                var file = ctx.Web.GetFileByServerRelativeUrl(Uri.EscapeUriString(itemUrl));
                ctx.Load(file,
                    f => f.ListItemAllFields,
                    f => f.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl);
                ctx.ExecuteQuery();
                return Uri.UnescapeDataString(
                    file.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl);
            }
        }

        // -------------------------------------------------------------------
        //  USER SEARCH
        // -------------------------------------------------------------------

        public List<UserModel> SearchUsers(string term)
        {
            var results = new List<UserModel>();
            try
            {
                using (var pc = GetPrincipalContext())
                using (var searcher = new PrincipalSearcher(new UserPrincipal(pc)))
                {
                    ((UserPrincipal)searcher.QueryFilter).SamAccountName = $"*{term}*";
                    foreach (System.DirectoryServices.AccountManagement.Principal p in searcher.FindAll().Take(15))
                    {
                        if (p is UserPrincipal up)
                            results.Add(new UserModel
                            {
                                DisplayName = up.DisplayName ?? up.SamAccountName,
                                Login       = up.SamAccountName,
                                Email       = up.EmailAddress ?? ""
                            });
                    }
                }
            }
            catch { }
            return results;
        }

        // -------------------------------------------------------------------
        //  INFRASTRUCTURE SETUP  (idempotent - safe to run multiple times)
        // -------------------------------------------------------------------

        public void SetupWorkflowInfrastructure()
        {
            using (var ctx = GetContext())
            {
                var web = ctx.Web;
                ctx.Load(web, w => w.Lists);
                ctx.ExecuteQuery();

                var wfConfig  = EnsureList(ctx, web, "WorkflowConfigs", 100);
                EnsureField(ctx, wfConfig, "LibraryUrl",        FieldType.Text);
                EnsureField(ctx, wfConfig, "Levels",            FieldType.Number);
                EnsureField(ctx, wfConfig, "IsActive",          FieldType.Boolean);
                EnsureField(ctx, wfConfig, "StagesJson",        FieldType.Note);
                EnsureField(ctx, wfConfig, "TriggerEvents",     FieldType.Note);
                EnsureField(ctx, wfConfig, "RejectionBehavior", FieldType.Text);
                EnsureField(ctx, wfConfig, "NotifyOnSubmit",    FieldType.Boolean);
                EnsureField(ctx, wfConfig, "NotifyOnApprove",   FieldType.Boolean);
                EnsureField(ctx, wfConfig, "NotifyOnReject",    FieldType.Boolean);
                EnsureField(ctx, wfConfig, "NotifyOnEscalate",  FieldType.Boolean);
                EnsureField(ctx, wfConfig, "NotifyOnDelegate",  FieldType.Boolean);
                EnsureField(ctx, wfConfig, "NotifyOnComplete",  FieldType.Boolean);

                var approvals = EnsureList(ctx, web, "ApprovalInstances", 100);
                EnsureField(ctx, approvals, "WorkflowRunId",    FieldType.Text);
                EnsureField(ctx, approvals, "ItemUrl",          FieldType.Text);
                EnsureField(ctx, approvals, "WorkflowId",       FieldType.Text);
                EnsureField(ctx, approvals, "WorkflowName",     FieldType.Text);
                EnsureField(ctx, approvals, "Stage",            FieldType.Number);
                EnsureField(ctx, approvals, "TotalLevels",      FieldType.Number);
                EnsureField(ctx, approvals, "ApprovalMode",     FieldType.Text);
                EnsureField(ctx, approvals, "Approver",         FieldType.Text);
                EnsureField(ctx, approvals, "OriginalApprover", FieldType.Text);
                EnsureField(ctx, approvals, "Status",           FieldType.Text);
                EnsureField(ctx, approvals, "SubmittedBy",      FieldType.Text);
                EnsureField(ctx, approvals, "SubmittedDate",    FieldType.DateTime);
                EnsureField(ctx, approvals, "ActionDate",       FieldType.DateTime);
                EnsureField(ctx, approvals, "DueDate",          FieldType.DateTime);
                EnsureField(ctx, approvals, "Comments",         FieldType.Note);
                EnsureField(ctx, approvals, "IsEscalated",      FieldType.Boolean);

                var audit = EnsureList(ctx, web, "WorkflowAuditLog", 100);
                EnsureField(ctx, audit, "WorkflowRunId", FieldType.Text);
                EnsureField(ctx, audit, "ItemUrl",       FieldType.Text);
                EnsureField(ctx, audit, "ItemName",      FieldType.Text);
                EnsureField(ctx, audit, "WorkflowName",  FieldType.Text);
                EnsureField(ctx, audit, "Actor",         FieldType.Text);
                EnsureField(ctx, audit, "Action",        FieldType.Text);
                EnsureField(ctx, audit, "Stage",         FieldType.Number);
                EnsureField(ctx, audit, "Comments",      FieldType.Note);
                EnsureField(ctx, audit, "Timestamp",     FieldType.DateTime);
            }
        }

        private List EnsureList(ClientContext ctx, Web web, string title, int template)
        {
            try
            {
                var existing = web.Lists.GetByTitle(title);
                ctx.Load(existing);
                ctx.ExecuteQuery();
                return existing;
            }
            catch
            {
                var newList = web.Lists.Add(new ListCreationInformation
                {
                    Title             = title,
                    TemplateType      = template,
                    QuickLaunchOption = QuickLaunchOptions.Off
                });
                ctx.ExecuteQuery();
                return newList;
            }
        }

        private void EnsureField(ClientContext ctx, List list, string name, FieldType type)
        {
            ctx.Load(list.Fields);
            ctx.ExecuteQuery();
            if (list.Fields.Any(f => f.InternalName == name)) return;
            list.Fields.AddFieldAsXml(
                $"<Field Type='{type}' Name='{name}' StaticName='{name}' DisplayName='{name}'/>",
                true, AddFieldOptions.DefaultValue);
            ctx.ExecuteQuery();
        }

        // -------------------------------------------------------------------
        //  PRIVATE ENGINE HELPERS
        // -------------------------------------------------------------------

        private bool IsStageComplete(ClientContext ctx, List list,
                                     string runId, int stage, string approvalMode)
        {
            var caml = new CamlQuery
            {
                ViewXml = $@"<View><Query><Where>
                    <And>
                      <Eq><FieldRef Name='WorkflowRunId'/><Value Type='Text'>{runId}</Value></Eq>
                      <Eq><FieldRef Name='Stage'/><Value Type='Number'>{stage}</Value></Eq>
                    </And>
                </Where></Query></View>"
            };
            var items = list.GetItems(caml);
            ctx.Load(items);
            ctx.ExecuteQuery();
            var statuses = new List<string>();
            foreach (var i in items) statuses.Add(i["Status"] != null ? i["Status"].ToString() : null);

            if (string.Equals(approvalMode, "All", StringComparison.OrdinalIgnoreCase))
                return statuses.Count > 0 &&
                       statuses.All(s => s == "Approved" || s == "Superseded");
            return statuses.Any(s => s == "Approved");
        }

        private void SupersedePendingInStage(ClientContext ctx, List list,
                                              string runId, int stage, string excludeId)
        {
            var caml = new CamlQuery
            {
                ViewXml = $@"<View><Query><Where>
                    <And>
                      <Eq><FieldRef Name='WorkflowRunId'/><Value Type='Text'>{runId}</Value></Eq>
                      <And>
                        <Eq><FieldRef Name='Stage'/><Value Type='Number'>{stage}</Value></Eq>
                        <Eq><FieldRef Name='Status'/><Value Type='Text'>Pending</Value></Eq>
                      </And>
                    </And>
                </Where></Query></View>"
            };
            var items = list.GetItems(caml);
            ctx.Load(items);
            ctx.ExecuteQuery();
            bool changed = false;
            foreach (var item in items)
            {
                if (item.Id.ToString() == excludeId) continue;
                item["Status"] = "Superseded";
                item.Update();
                changed = true;
            }
            if (changed) ctx.ExecuteQuery();
        }

        private void ActivateStage(ClientContext ctx, List list, string runId, int stage)
        {
            var caml = new CamlQuery
            {
                ViewXml = $@"<View><Query><Where>
                    <And>
                      <Eq><FieldRef Name='WorkflowRunId'/><Value Type='Text'>{runId}</Value></Eq>
                      <And>
                        <Eq><FieldRef Name='Stage'/><Value Type='Number'>{stage}</Value></Eq>
                        <Eq><FieldRef Name='Status'/><Value Type='Text'>Waiting</Value></Eq>
                      </And>
                    </And>
                </Where></Query></View>"
            };
            var items = list.GetItems(caml);
            ctx.Load(items);
            ctx.ExecuteQuery();

            int dueInDays = 3;
            if (items.Count > 0)
            {
                string wfId = items[0]["WorkflowId"]?.ToString();
                if (int.TryParse(wfId, out int wfIdInt))
                {
                    var sm = GetWorkflows().FirstOrDefault(w => w.Id == wfIdInt)
                                          ?.Stages?.FirstOrDefault(s => s.Level == stage);
                    if (sm != null) dueInDays = sm.DueInDays;
                }
            }

            DateTime? due = dueInDays > 0 ? (DateTime?)DateTime.Now.AddDays(dueInDays) : null;
            foreach (var item in items)
            {
                item["Status"]      = "Pending";
                item["DueDate"]     = due.HasValue ? (object)due.Value : null;
                item["IsEscalated"] = false;
                item.Update();
            }
            if (items.Count > 0) ctx.ExecuteQuery();
        }

        private void CancelStageInstances(ClientContext ctx, List list,
                                          string runId, int stage)
        {
            var caml = new CamlQuery
            {
                ViewXml = $@"<View><Query><Where>
                    <And>
                      <Eq><FieldRef Name='WorkflowRunId'/><Value Type='Text'>{runId}</Value></Eq>
                      <Eq><FieldRef Name='Stage'/><Value Type='Number'>{stage}</Value></Eq>
                    </And>
                </Where></Query></View>"
            };
            var items = list.GetItems(caml);
            ctx.Load(items);
            ctx.ExecuteQuery();
            foreach (var item in items)
            {
                string s = item["Status"]?.ToString();
                if (s == "Pending" || s == "Waiting") { item["Status"] = "Cancelled"; item.Update(); }
            }
            ctx.ExecuteQuery();
        }

        private void CancelAllPending(ClientContext ctx, List list, string runId)
        {
            var caml = new CamlQuery
            {
                ViewXml = $@"<View><Query><Where>
                    <Eq><FieldRef Name='WorkflowRunId'/><Value Type='Text'>{runId}</Value></Eq>
                </Where></Query></View>"
            };
            var items = list.GetItems(caml);
            ctx.Load(items);
            ctx.ExecuteQuery();
            bool any = false;
            foreach (var item in items)
            {
                string s = item["Status"]?.ToString();
                if (s == "Pending" || s == "Waiting") { item["Status"] = "Cancelled"; item.Update(); any = true; }
            }
            if (any) ctx.ExecuteQuery();
        }

        // Type-safe SP field readers
        private static int SafeInt(object val)
        {
            if (val == null) return 0;
            if (double.TryParse(val.ToString(), out double d)) return Convert.ToInt32(d);
            return 0;
        }

        private static bool SafeBool(object val)
        {
            if (val == null) return false;
            return string.Equals(val.ToString(), "true", StringComparison.OrdinalIgnoreCase)
                || val.ToString() == "1";
        }

        private static DateTime? SafeDate(object val)
        {
            if (val == null) return null;
            return DateTime.TryParse(val.ToString(), out DateTime dt) ? (DateTime?)dt : null;
        }

        private static T DeserializeJson<T>(string json) where T : class
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            try { return JsonConvert.DeserializeObject<T>(json); }
            catch { return null; }
        }

        private string NormalizeUrl(string url) =>
            string.IsNullOrWhiteSpace(url)
                ? string.Empty
                : Uri.UnescapeDataString(url).Trim().ToLowerInvariant();

        private string GetSiteBaseUrl()
        {
            try
            {
                var req = HttpContext.Current?.Request;
                if (req != null) return $"{req.Url.Scheme}://{req.Url.Authority}";
            }
            catch { }
            return siteUrl;
        }

        private string ResolveEmail(string usernameOrEmail)
        {
            if (string.IsNullOrEmpty(usernameOrEmail)) return "";
            if (usernameOrEmail.Contains("@")) return usernameOrEmail;
            try
            {
                using (var pc = GetPrincipalContext())
                {
                    var up = UserPrincipal.FindByIdentity(pc, usernameOrEmail);
                    if (!string.IsNullOrEmpty(up?.EmailAddress)) return up.EmailAddress;
                }
            }
            catch { }
            return usernameOrEmail;
        }
    }
}
