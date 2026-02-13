using System;
using System.Configuration;
using System.Net;
using System.Web.Mvc;
using KCAU_SharePoint.Models;
using Microsoft.SharePoint.Client;

namespace KCAU_SharePoint.Controllers
{
    public class AccountController : Controller
    {
		Helper helper;

		public AccountController()
		{
            helper = new Helper();
        }

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

                // ✅ LOGIN SUCCESS: store in session
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
