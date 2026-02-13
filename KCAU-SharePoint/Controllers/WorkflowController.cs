using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using KCAU_SharePoint.Models;

public class WorkflowController : Controller
{
    private Helper helper;
    private static List<WorkflowInstance> WorkflowInstances = new List<WorkflowInstance>();
    private static List<WorkflowApprovalHistory> ApprovalHistories = new List<WorkflowApprovalHistory>();

    public WorkflowController()
    {
        helper = new Helper();
    }

    public ActionResult Index()
    {
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


    // Dashboard for approvers
    public ActionResult MyApprovals()
    {
        var user = User.Identity.Name;
        var pending = new List<WorkflowInstance>();

        foreach (var instance in WorkflowInstances.Where(i => i.Status == "Pending"))
        {
            var workflow = helper.GetWorkflows().FirstOrDefault(w => w.Id == instance.WorkflowId);
            var stage = workflow?.Stages.FirstOrDefault(s => s.Level == instance.CurrentLevel);
            if (stage != null && stage.Approvers.Contains(user, StringComparer.OrdinalIgnoreCase))
            {
                pending.Add(instance);
            }
        }

        return View(pending);
    }

    [HttpPost]
    public JsonResult Approve(int instanceId, string comments = "")
    {
        try
        {
            var user = User.Identity.Name;
            var instance = WorkflowInstances.FirstOrDefault(i => i.Id == instanceId);
            if (instance == null) return Json(new { success = false, message = "Instance not found" });

            var workflow = helper.GetWorkflows().FirstOrDefault(w => w.Id == instance.WorkflowId);
            var stage = workflow.Stages.FirstOrDefault(s => s.Level == instance.CurrentLevel);

            if (!stage.Approvers.Contains(user, StringComparer.OrdinalIgnoreCase))
                return Json(new { success = false, message = "Not authorized" });

            // Log approval
            ApprovalHistories.Add(new WorkflowApprovalHistory
            {
                Id = ApprovalHistories.Count + 1,
                WorkflowInstanceId = instance.Id,
                Level = instance.CurrentLevel,
                Approver = user,
                Action = "Approved",
                Comments = comments,
                ActionDate = DateTime.Now
            });

            // Move to next level
            if (instance.CurrentLevel < instance.TotalLevels)
            {
                instance.CurrentLevel++;
            }
            else
            {
                instance.Status = "Approved";
                instance.CompletedDate = DateTime.Now;
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
            var instance = WorkflowInstances.FirstOrDefault(i => i.Id == instanceId);
            if (instance == null) return Json(new { success = false, message = "Instance not found" });

            var workflow = helper.GetWorkflows().FirstOrDefault(w => w.Id == instance.WorkflowId);
            var stage = workflow.Stages.FirstOrDefault(s => s.Level == instance.CurrentLevel);

            if (!stage.Approvers.Contains(user, StringComparer.OrdinalIgnoreCase))
                return Json(new { success = false, message = "Not authorized" });

            // Log rejection
            ApprovalHistories.Add(new WorkflowApprovalHistory
            {
                Id = ApprovalHistories.Count + 1,
                WorkflowInstanceId = instance.Id,
                Level = instance.CurrentLevel,
                Approver = user,
                Action = "Rejected",
                Comments = comments,
                ActionDate = DateTime.Now
            });

            instance.Status = "Rejected";
            instance.CompletedDate = DateTime.Now;

            return Json(new { success = true });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = ex.Message });
        }
    }

}
