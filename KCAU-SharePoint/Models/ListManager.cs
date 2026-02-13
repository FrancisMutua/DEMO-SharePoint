using System;
using Microsoft.SharePoint.Client;

namespace KCAU_SharePoint.Models
{
    /// <summary>
    /// Manages creation and setup of workflow-related SharePoint lists
    /// </summary>
    public class ListManager
    {
        private readonly Helper helper;

        public ListManager()
        {
            helper = new Helper();
        }

        /// <summary>
        /// Creates all required lists for the workflow system
        /// Call this once during initial setup
        /// </summary>
        public void CreateWorkflowLists()
        {
            using (var ctx = helper.GetContext())
            {
                // Create WorkflowConfigs list
                CreateWorkflowConfigsList(ctx);

                // Create WorkflowLevels list
                CreateWorkflowLevelsList(ctx);

                // Create WorkflowInstances list
                CreateWorkflowInstancesList(ctx);

                // Create WorkflowHistory list
                CreateWorkflowHistoryList(ctx);
            }
        }

        private void CreateWorkflowConfigsList(ClientContext ctx)
        {
            var lists = ctx.Web.Lists;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            // Check if list already exists
            foreach (var list in lists)
            {
                if (list.Title == "WorkflowConfigs")
                {
                    Console.WriteLine("WorkflowConfigs list already exists");
                    return;
                }
            }

            // Create list
            var listCreationInfo = new ListCreationInformation
            {
                Title = "WorkflowConfigs",
                Description = "Stores workflow configuration settings",
                TemplateType = (int)ListTemplateType.GenericList
            };

            var newList = ctx.Web.Lists.Add(listCreationInfo);
            ctx.Load(newList);
            ctx.ExecuteQuery();

            // Add custom columns
            newList.Fields.AddFieldAsXml(
                "<Field Type='Text' DisplayName='LibraryUrl' Required='TRUE' MaxLength='255' Name='LibraryUrl'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Number' DisplayName='Levels' Required='TRUE' Name='Levels'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Boolean' DisplayName='IsActive' Required='FALSE' Name='IsActive'><Default>1</Default></Field>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            ctx.ExecuteQuery();

            Console.WriteLine("WorkflowConfigs list created successfully");
        }

        private void CreateWorkflowLevelsList(ClientContext ctx)
        {
            var lists = ctx.Web.Lists;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            // Check if list already exists
            foreach (var list in lists)
            {
                if (list.Title == "WorkflowLevels")
                {
                    Console.WriteLine("WorkflowLevels list already exists");
                    return;
                }
            }

            // Create list
            var listCreationInfo = new ListCreationInformation
            {
                Title = "WorkflowLevels",
                Description = "Stores approval levels and approvers for each workflow",
                TemplateType = (int)ListTemplateType.GenericList
            };

            var newList = ctx.Web.Lists.Add(listCreationInfo);
            ctx.Load(newList);
            ctx.ExecuteQuery();

            // Add custom columns
            newList.Fields.AddFieldAsXml(
                "<Field Type='Number' DisplayName='WorkflowId' Required='TRUE' Name='WorkflowId'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Number' DisplayName='LevelNo' Required='TRUE' Name='LevelNo'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='UserMulti' DisplayName='Approvers' Required='TRUE' Name='Approvers' Mult='TRUE' UserSelectionMode='PeopleOnly'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            ctx.ExecuteQuery();

            Console.WriteLine("WorkflowLevels list created successfully");
        }

        private void CreateWorkflowInstancesList(ClientContext ctx)
        {
            var lists = ctx.Web.Lists;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            // Check if list already exists
            foreach (var list in lists)
            {
                if (list.Title == "WorkflowInstances")
                {
                    Console.WriteLine("WorkflowInstances list already exists");
                    return;
                }
            }

            // Create list
            var listCreationInfo = new ListCreationInformation
            {
                Title = "WorkflowInstances",
                Description = "Tracks active and completed workflow instances",
                TemplateType = (int)ListTemplateType.GenericList
            };

            var newList = ctx.Web.Lists.Add(listCreationInfo);
            ctx.Load(newList);
            ctx.ExecuteQuery();

            // Add custom columns
            newList.Fields.AddFieldAsXml(
                "<Field Type='Number' DisplayName='WorkflowId' Required='TRUE' Name='WorkflowId'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Text' DisplayName='ItemUrl' Required='TRUE' MaxLength='500' Name='ItemUrl'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Number' DisplayName='CurrentLevel' Required='TRUE' Name='CurrentLevel'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Choice' DisplayName='Status' Required='TRUE' Format='Dropdown' Name='Status'>" +
                "<CHOICES><CHOICE>Pending</CHOICE><CHOICE>Approved</CHOICE><CHOICE>Rejected</CHOICE></CHOICES>" +
                "<Default>Pending</Default></Field>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='DateTime' DisplayName='StartedDate' Required='TRUE' Format='DateTime' Name='StartedDate'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='DateTime' DisplayName='CompletedDate' Required='FALSE' Format='DateTime' Name='CompletedDate'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            ctx.ExecuteQuery();

            Console.WriteLine("WorkflowInstances list created successfully");
        }

        private void CreateWorkflowHistoryList(ClientContext ctx)
        {
            var lists = ctx.Web.Lists;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            // Check if list already exists
            foreach (var list in lists)
            {
                if (list.Title == "WorkflowHistory")
                {
                    Console.WriteLine("WorkflowHistory list already exists");
                    return;
                }
            }

            // Create list
            var listCreationInfo = new ListCreationInformation
            {
                Title = "WorkflowHistory",
                Description = "Logs all approval and rejection actions",
                TemplateType = (int)ListTemplateType.GenericList
            };

            var newList = ctx.Web.Lists.Add(listCreationInfo);
            ctx.Load(newList);
            ctx.ExecuteQuery();

            // Add custom columns
            newList.Fields.AddFieldAsXml(
                "<Field Type='Number' DisplayName='InstanceId' Required='TRUE' Name='InstanceId'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Number' DisplayName='Level' Required='TRUE' Name='Level'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Text' DisplayName='Approver' Required='TRUE' MaxLength='255' Name='Approver'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Choice' DisplayName='Action' Required='TRUE' Format='Dropdown' Name='Action'>" +
                "<CHOICES><CHOICE>Approved</CHOICE><CHOICE>Rejected</CHOICE></CHOICES></Field>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='Note' DisplayName='Comments' Required='FALSE' NumLines='6' Name='Comments'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            newList.Fields.AddFieldAsXml(
                "<Field Type='DateTime' DisplayName='ActionDate' Required='TRUE' Format='DateTime' Name='ActionDate'/>",
                true,
                AddFieldOptions.AddFieldInternalNameHint);

            ctx.ExecuteQuery();

            Console.WriteLine("WorkflowHistory list created successfully");
        }

        /// <summary>
        /// Checks if all required lists exist
        /// </summary>
        public bool ValidateListsExist()
        {
            using (var ctx = helper.GetContext())
            {
                var lists = ctx.Web.Lists;
                ctx.Load(lists);
                ctx.ExecuteQuery();

                bool hasWorkflowConfigs = false;
                bool hasWorkflowLevels = false;
                bool hasWorkflowInstances = false;
                bool hasWorkflowHistory = false;

                foreach (var list in lists)
                {
                    if (list.Title == "WorkflowConfigs") hasWorkflowConfigs = true;
                    if (list.Title == "WorkflowLevels") hasWorkflowLevels = true;
                    if (list.Title == "WorkflowInstances") hasWorkflowInstances = true;
                    if (list.Title == "WorkflowHistory") hasWorkflowHistory = true;
                }

                return hasWorkflowConfigs && hasWorkflowLevels && hasWorkflowInstances && hasWorkflowHistory;
            }
        }
    }
}