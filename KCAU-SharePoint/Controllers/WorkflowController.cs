using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using KCAU_SharePoint.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using Newtonsoft.Json;
using static KCAU_SharePoint.Models.Helper;
using WorkflowInstance = KCAU_SharePoint.Models.WorkflowInstance;

[SessionAuthorize]
public class WorkflowController : Controller
{
    private Helper helper;

    public WorkflowController()
    {
        helper = new Helper();
    }

    public ActionResult Index()
    {
        var libraries = helper.GetDocumentLibraries();
        ViewBag.Libraries = libraries;
        var workflows = helper.GetWorkflows();

        var vm = new WorkflowConfigVM
        {
            Libraries = helper.GetDocumentLibraries(),
            ExistingWorkflows = workflows,
            ApprovalLevels = Enumerable.Range(1, 10)
                .Select(x => new SelectListItem
                {
                    Text = x + " Levels",
                    Value = x.ToString()
                })
                .ToList()
        };

        return View(vm);
    }



    [HttpPost]
    [ValidateInput(false)]
    public JsonResult CreateWorkflow(CreateWorkflowRequest model)
    {
        try
        {
            if (model == null || string.IsNullOrEmpty(model.Name) ||
                model.Stages == null || model.Stages.Count != model.Levels)
            {
                return Json(new { success = false, message = "Invalid workflow data." });
            }

            helper.CreateWorkflow(model);

            return Json(new { success = true });
        }catch(Exception ex)
        {
            return Json(new { success = false,message=ex.Message });
        }
    }


    [HttpPost]

    [ValidateInput(false)]
    public JsonResult DeleteWorkflow(int id)
    {
        try
        {
            helper.DeleteWorkflow(id); // create this method
            return Json(new { success = true });
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
            var librayUrl= helper.GetLibraryUrlFromItem(itemUrl);
            var workflow = helper.GetWorkflowForLibrary(librayUrl);

            if (workflow == null)
                return Json(new { success = false, message = "No workflow configured." });

            string username = HttpContext.Session["Username"]?.ToString() ?? "UnknownUser";

            helper.CreateWorkflowInstance(itemUrl, itemName, username);

            return Json(new { success = true });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = ex.Message });
        }
    }


    // Dashboard for approvers
    public ActionResult MyApprovals()
    {
        var user = HttpContext.Session["Username"]?.ToString() ?? "UnknownUser";
        var pending = new List<WorkflowInstance>();

        // Get all workflow instances with Status = "Pending"
        var allPendingInstances = helper.GetWorkflowInstances()
                                        .Where(i => i.Status.Equals("Pending", StringComparison.OrdinalIgnoreCase))
                                        .ToList();

        foreach (var instance in allPendingInstances)
        {
            // Check if the logged-in user is an approver for this stage
            if (instance.Approver != null && instance.Approver.Equals(user, StringComparison.OrdinalIgnoreCase))
            {
                pending.Add(instance);
            }
        }

        return View(pending);
    }


    [HttpPost]
    public JsonResult Approve(int instanceId, string itemUrl, string comments = "")
    {
        try
        {
            var user = User.Identity.Name;

            // Get all workflow instances
            var allInstances = helper.GetWorkflowInstances();

            // Find the current instance by ID
            var instance = allInstances.FirstOrDefault(i => i.Id == instanceId.ToString());
            if (instance == null)
                return Json(new { success = false, message = "Workflow instance not found." });

            // Find the pending instance for this user
            var pendingInstance = allInstances
                .FirstOrDefault(i => i.ItemUrl == instance.ItemUrl && i.Status == "Pending" && i.Approver == user);

            if (pendingInstance == null)
                return Json(new { success = false, message = "No pending approval found for this user on this document." });

            // Mark current instance as Approved
            pendingInstance.Status = "Approved";
            pendingInstance.CompletedDate = DateTime.Now;
            pendingInstance.Comments = comments;

            
            // Move next stage's first approver to Pending
            string libraryUrl = helper.GetLibraryUrlFromItem(itemUrl);
            var workflow = helper.GetWorkflowForLibrary(libraryUrl);

            if (workflow != null)
            {
                int nextLevel = int.Parse(pendingInstance.CurrentLevel) + 1;
                var nextStage = workflow.Stages.FirstOrDefault(s => s.Level == nextLevel);

                if (nextStage != null && nextStage.Approvers.Any())
                {
                    var nextInstance = allInstances.FirstOrDefault(i =>
                        i.ItemUrl == pendingInstance.ItemUrl &&
                        i.CurrentLevel == nextLevel.ToString() &&
                        i.Approver == nextStage.Approvers[0] &&
                        i.Status == "Created");

                    if (nextInstance != null)
                    {
                        nextInstance.Status = "Pending";
                    }
                }
            }

            return Json(new { success = true });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = ex.Message });
        }
    }



    [HttpPost]
    public JsonResult Reject(int instanceId, string comments = "")
    {
        try
        {
            var user = User.Identity.Name;

            // Get the workflow instance
            var instance = helper.GetWorkflowInstances().FirstOrDefault(i => i.Id.Equals( instanceId));
            if (instance == null)
                return Json(new { success = false, message = "Instance not found" });

            // Get the workflow configuration
            var workflow = helper.GetWorkflows().FirstOrDefault(w => w.Id.Equals( instance.WorkflowId));
            if (workflow == null)
                return Json(new { success = false, message = "Workflow configuration not found" });

            // Get current stage
            var stage = workflow.Stages.FirstOrDefault(s => s.Level .Equals(instance.CurrentLevel));
            if (stage == null)
                return Json(new { success = false, message = "Current stage not found" });

            // Check if user is authorized to reject
            if (!stage.Approvers.Contains(user, StringComparer.OrdinalIgnoreCase))
                return Json(new { success = false, message = "Not authorized to reject this instance" });

            // Update SharePoint list item (if using CSOM)
            using (var ctx = helper.GetContext())
            {
                var list = ctx.Web.Lists.GetByTitle("ApprovalInstances");
                var spItem = list.GetItemById(instance.Id);

                spItem["Status"] = "Rejected";
                spItem["CompletedDate"] = DateTime.Now;
                ctx.ExecuteQuery();
            }

            return Json(new { success = true });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = ex.Message });
        }
    }


}
