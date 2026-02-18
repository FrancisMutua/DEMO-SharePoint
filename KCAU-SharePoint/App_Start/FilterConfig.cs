using System.Web;
using System.Web.Mvc;
using KCAU_SharePoint.Models;

namespace KCAU_SharePoint
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
