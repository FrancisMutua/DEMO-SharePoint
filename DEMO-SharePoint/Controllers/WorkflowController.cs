using System;
using System.Linq;
using System.Web.Mvc;
using Demo_SharePoint.Services.Implementations;
using DEMO_SharePoint.Models;

namespace DEMO_SharePoint.Controllers
{
    [SessionAuthorize]
    public class WorkflowController : BaseController
    {
        private readonly Helper helper;

        public WorkflowController()
        {
            helper = new Helper();
        }

        private string CurrentUser =>
            HttpContext.Session["Username"]?.ToString() ?? "UnknownUser";

        private bool IsAdmin =>
            CurrentUser.Equals("Administrator", StringComparison.OrdinalIgnoreCase);

        // -------------------------------------------------------------------
        //  WORKFLOW MANAGEMENT  (admin)
        // -------------------------------------------------------------------

        public ActionResult Index()
        {
            try
            {
                var workflows = helper.GetWorkflows();
                var vm = new WorkflowConfigVM
                {
                    Libraries         = helper.GetDocumentLibraries(),
                    ExistingWorkflows = workflows,
                    ApprovalLevels    = Enumerable.Range(1, 10)
                        .Select(x => new SelectListItem
                        {
                            Text  = $"{x} Level{(x > 1 ? "s" : "")}",
                            Value = x.ToString()
                        }).ToList()
                };
                return View(vm);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorTitle = "Workflow Configuration Unavailable";
                ViewBag.ErrorMessage = "Workflow settings could not be loaded. Please check your SharePoint connection and try again.";
                ViewBag.ErrorDetail = ex.Message;
                //return View("Error");
                return View(new WorkflowConfigVM());
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public JsonResult CreateWorkflow(CreateWorkflowRequest model)
        {
            if (!IsAdmin)
                return Json(new { success = false, message = "Only Administrators can manage workflows." });
            try
            {
                if (model == null || string.IsNullOrEmpty(model.Name) ||
                    string.IsNullOrEmpty(model.LibraryUrl) ||
                    model.Stages == null || model.Stages.Count != model.Levels)
                    return Json(new { success = false, message = "Invalid workflow data. Ensure all levels are configured." });

                helper.CreateWorkflow(model);
                return Json(new { success = true });
            }
            catch (Exception ex) { return Json(new { success = false, message = ex.Message }); }
        }

        [HttpPost]
        public JsonResult DeleteWorkflow(int id)
        {
            if (!IsAdmin)
                return Json(new { success = false, message = "Only Administrators can delete workflows." });
            try { helper.DeleteWorkflow(id); return Json(new { success = true }); }
            catch (Exception ex) { return Json(new { success = false, message = ex.Message }); }
        }

        [HttpPost]
        public JsonResult ToggleWorkflow(int id, bool isActive)
        {
            if (!IsAdmin)
                return Json(new { success = false, message = "Only Administrators can toggle workflows." });
            try { helper.ToggleWorkflow(id, isActive); return Json(new { success = true }); }
            catch (Exception ex) { return Json(new { success = false, message = ex.Message }); }
        }

        // -------------------------------------------------------------------
        //  SUBMIT FOR APPROVAL
        // -------------------------------------------------------------------

        [HttpPost]
        public JsonResult SubmitForApproval(string itemUrl, string itemName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(itemUrl))
                    return Json(new { success = false, message = "Item URL is required." });

                string runId = helper.CreateWorkflowRun(
                    itemUrl, itemName ?? "Untitled", CurrentUser, "Manual");
                return Json(new { success = true, workflowRunId = runId });
            }
            catch (Exception ex) { return Json(new { success = false, message = ex.Message }); }
        }

        // -------------------------------------------------------------------
        //  APPROVER DASHBOARD - My Approvals
        // -------------------------------------------------------------------

        public ActionResult MyApprovals()
        {
            try
            {
                var items = helper.GetPendingApprovalsForUser(CurrentUser);
                return View(items);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorTitle = "Approvals Unavailable";
                ViewBag.ErrorMessage = "Your pending approvals could not be loaded. Please check your SharePoint connection and try again.";
                ViewBag.ErrorDetail = ex.Message;
                return View("Error");
            }
        }

        // -------------------------------------------------------------------
        //  SUBMITTER DASHBOARD - Submitted Approvals
        // -------------------------------------------------------------------

        public ActionResult SubmittedApprovals()
        {
            try
            {
                var all = helper.GetSubmittedByUser(CurrentUser);

                // Collapse to one representative row per workflow run
                var grouped = all
                    .GroupBy(i => i.WorkflowRunId)
                    .Select(g =>
                    {
                        var rep = g.OrderBy(i => int.TryParse(i.CurrentLevel, out int l) ? l : 0)
                                   .FirstOrDefault(i => i.Status == "Pending")
                               ?? g.OrderByDescending(i => i.SubmittedDate).First();
                        rep.TotalLevels = g.Max(i => i.TotalLevels);
                        return rep;
                    })
                    .OrderByDescending(i => i.SubmittedDate)
                    .ToList();

                return View(grouped);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorTitle = "Submitted Approvals Unavailable";
                ViewBag.ErrorMessage = "Your submitted documents could not be loaded. Please check your SharePoint connection and try again.";
                ViewBag.ErrorDetail = ex.Message;
                return View("Error");
            }
        }

        // -------------------------------------------------------------------
        //  APPROVE
        // -------------------------------------------------------------------

        [HttpPost]
        public JsonResult Approve(ApproveRequest req)
        {
            try
            {
                if (req == null || string.IsNullOrEmpty(req.InstanceId))
                    return Json(new { success = false, message = "Invalid request." });
                helper.ApproveInstance(req.InstanceId, CurrentUser, req.Comments);
                return Json(new { success = true });
            }
            catch (UnauthorizedAccessException ex)
            { return Json(new { success = false, message = ex.Message }); }
            catch (Exception ex)
            { return Json(new { success = false, message = ex.Message }); }
        }

        // -------------------------------------------------------------------
        //  REJECT
        // -------------------------------------------------------------------

        [HttpPost]
        public JsonResult Reject(RejectRequest req)
        {
            try
            {
                if (req == null || string.IsNullOrEmpty(req.InstanceId))
                    return Json(new { success = false, message = "Invalid request." });
                helper.RejectInstance(req.InstanceId, CurrentUser, req.Comments);
                return Json(new { success = true });
            }
            catch (UnauthorizedAccessException ex)
            { return Json(new { success = false, message = ex.Message }); }
            catch (Exception ex)
            { return Json(new { success = false, message = ex.Message }); }
        }

        // -------------------------------------------------------------------
        //  DELEGATE
        // -------------------------------------------------------------------

        [HttpPost]
        public JsonResult Delegate(DelegateRequest req)
        {
            try
            {
                if (req == null || string.IsNullOrEmpty(req.InstanceId) ||
                    string.IsNullOrEmpty(req.ToUser))
                    return Json(new { success = false, message = "Invalid request." });
                helper.DelegateInstance(req.InstanceId, CurrentUser, req.ToUser, req.Reason);
                return Json(new { success = true });
            }
            catch (UnauthorizedAccessException ex)
            { return Json(new { success = false, message = ex.Message }); }
            catch (Exception ex)
            { return Json(new { success = false, message = ex.Message }); }
        }

        // -------------------------------------------------------------------
        //  RECALL
        // -------------------------------------------------------------------

        [HttpPost]
        public JsonResult Recall(RecallRequest req)
        {
            try
            {
                if (req == null || string.IsNullOrEmpty(req.WorkflowRunId))
                    return Json(new { success = false, message = "Invalid request." });
                helper.RecallWorkflowRun(req.WorkflowRunId, CurrentUser);
                return Json(new { success = true });
            }
            catch (UnauthorizedAccessException ex)
            { return Json(new { success = false, message = ex.Message }); }
            catch (Exception ex)
            { return Json(new { success = false, message = ex.Message }); }
        }

        // -------------------------------------------------------------------
        //  AUDIT LOG
        // -------------------------------------------------------------------

        public ActionResult AuditLog(string itemUrl)
        {
            try
            {
                if (string.IsNullOrEmpty(itemUrl))
                    return RedirectToAction("MyApprovals");
                ViewBag.ItemUrl = itemUrl;
                return View(helper.GetAuditLog(itemUrl));
            }
            catch (Exception ex)
            {
                ViewBag.ErrorTitle = "Audit Log Unavailable";
                ViewBag.ErrorMessage = "The audit log for this document could not be loaded. Please check your SharePoint connection and try again.";
                ViewBag.ErrorDetail = ex.Message;
                return View("Error");
            }
        }

        // -------------------------------------------------------------------
        //  ESCALATION  (admin / scheduler)
        // -------------------------------------------------------------------

        [HttpPost]
        public JsonResult RunEscalation()
        {
            if (!IsAdmin) return Json(new { success = false, message = "Admin only." });
            try
            {
                int count = helper.EscalateOverdueInstances();
                return Json(new { success = true, escalated = count });
            }
            catch (Exception ex) { return Json(new { success = false, message = ex.Message }); }
        }

        // -------------------------------------------------------------------
        //  INFRASTRUCTURE SETUP  (admin one-time)
        // -------------------------------------------------------------------

        [HttpPost]
        public JsonResult SetupInfrastructure()
        {
            if (!IsAdmin) return Json(new { success = false, message = "Admin only." });
            try
            {
                helper.SetupWorkflowInfrastructure();
                return Json(new { success = true, message = "SharePoint lists provisioned successfully." });
            }
            catch (Exception ex) { return Json(new { success = false, message = ex.Message }); }
        }

        // -------------------------------------------------------------------
        //  BADGE COUNT  (nav pending-item count)
        // -------------------------------------------------------------------

        public JsonResult GetPendingCount()
        {
            try
            {
                int c = helper.GetPendingApprovalsForUser(CurrentUser).Count;
                return Json(new { count = c }, JsonRequestBehavior.AllowGet);
            }
            catch { return Json(new { count = 0 }, JsonRequestBehavior.AllowGet); }
        }
    }
}
