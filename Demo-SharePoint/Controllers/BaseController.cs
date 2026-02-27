
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DEMO_SharePoint.Models;

namespace Demo_SharePoint.Services.Implementations
{
    public class BaseController : Controller
    {
        private readonly Helper _helper = new Helper();

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            ViewBag.Libraries = _helper.GetDocumentLibraries();
            base.OnActionExecuting(filterContext);
        }
    }
}
