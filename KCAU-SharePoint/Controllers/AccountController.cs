using System;
using System.Net;
using System.Web.Mvc;
using KCAU_SharePoint.Models;
using Microsoft.SharePoint.Client;

namespace KCAU_SharePoint.Controllers
{
    public class AccountController : Controller
    {
        // Your SharePoint site URL
        private string siteUrl = "http://41.89.240.139/sites/edms";

        public ActionResult Login()
        {
            return View(new UserModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Login(UserModel model)
        {
            if (!ModelState.IsValid)
                return View(model);

            // Authenticate via SharePoint
            if (ValidateUserSharePoint(model.Username, model.Password))
            {
                // Store username in session
                Session["Username"] = model.Username;

                return RedirectToAction("Index", "Home");
            }
            else
            {
                ModelState.AddModelError("", "Invalid username or password.");
                return View(model);
            }
        }

        private bool ValidateUserSharePoint(string username, string password)
        {
            try
            {
                using (var context = new ClientContext(siteUrl))
                {
                    // Domain name (Active Directory domain)
                    string domain = "UOEMDOMAIN"; 

                    // Use NetworkCredential for Windows Authentication
                    context.Credentials = new NetworkCredential(username, password, domain);

                    // Pre-authenticate to avoid multiple round trips
                    context.ExecutingWebRequest += (sender, e) =>
                    {
                        e.WebRequestExecutor.WebRequest.PreAuthenticate = true;
                    };

                    // Attempt to load a simple property to validate login
                    Web web = context.Web;
                    context.Load(web, w => w.Title);
                    context.ExecuteQuery(); 

                    return true; // Login successful
                }
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public ActionResult Logout()
        {
            Session.Clear();
            Session.Abandon();
            return RedirectToAction("Login");
        }
    }
}
