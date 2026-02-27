using System;
using System.Configuration;
using System.Net;
using System.Web.Mvc;
using DEMO_SharePoint.Models;
using Microsoft.SharePoint.Client;

namespace DEMO_SharePoint.Controllers
{
    public class AccountController : Controller
    {
		Helper helper;

		public AccountController()
		{
            helper = new Helper();
        }
        [AllowAnonymous]
        public ActionResult Login()
        {
            return View(new UserModel());
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult Login(UserModel model)
        {
            if (!ModelState.IsValid)
                return View(model);

            var helper = new Helper();

            try
            {
                // Pass username/password directly here
                using (var context = helper.GetContext(model.Username, model.Password))
                {
                    Web web = context.Web;
                    context.Load(web, w => w.Title);
                    context.ExecuteQuery(); // test credentials
                }

                // Login success: store in session
                Session["Username"] = model.Username;
                Session["password"] = model.Password;

                return RedirectToAction("Index", "Home");
            }
            catch (ClientRequestException)
            {
                ModelState.AddModelError("", "Unable to connect to SharePoint. Please try again later.");
            }
            catch (ServerUnauthorizedAccessException)
            {
                ModelState.AddModelError("", "Invalid username or password.");
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("", "Login failed: " + ex.Message);
            }

            return View(model);
        }

        public ActionResult Logout()
        {
            Session.Clear();
            Session.Abandon();
            return RedirectToAction("Login");
        }
    }
}
