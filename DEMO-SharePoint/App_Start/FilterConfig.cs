using System.Web;
using System.Web.Mvc;
using DEMO_SharePoint.Models;

namespace DEMO_SharePoint
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
            filters.Add(new SessionAuthorizeAttribute());
        }
    }
}
