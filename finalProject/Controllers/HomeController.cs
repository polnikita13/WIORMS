using finalProject.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace finalProject.Controllers
{
    //
    public class HomeController : Controller
    {
        private WIORMSDBEntities db = new WIORMSDBEntities();

        public ActionResult Index()
        {
            return View();
        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        //Admin Login
        public ActionResult ALogin()
        {
            ViewBag.msg = "";

            return View();
        }
        [HttpPost]
        public ActionResult ALogin(AdminLogin adminLogin)
        {

            AdminLogin ad = db.AdminLogins.FirstOrDefault(m => m.UserName == adminLogin.UserName && m.Password == adminLogin.Password);
            if (ad != null)
            {
                Session["logged"] = true;
                Session["user"] = ad.UserName;
                Session["role"] = "Admin";

                return RedirectToAction("Index");
            }

            ViewBag.msg = "Login Failed try again...";
            return View();
        }
        //Logout for all users
        public ActionResult Logout()
        {
            Session["logged"] = null;
            Session["user"] = null;
            Session["role"] = null;

            return RedirectToAction("Index", "Home");
        }

        //Member Login
        public ActionResult MemberLogin()
        {
            ViewBag.msg = "";

            return View();
        }
        [HttpPost]
        public ActionResult MemberLogin(String EmailId, String Password)
        {
            UserRegistration ad = db.UserRegistrations.FirstOrDefault(m => m.EmailId == EmailId && m.Password == Password);
            if (ad != null)
            {
                Session["logged"] = true;
                Session["user"] = ad.EmailId;
                Session["role"] = ad.Role;
                return RedirectToAction("Index");
            }
            ViewBag.msg = "Login Failed try again...";
            return View();
        }

    }
}