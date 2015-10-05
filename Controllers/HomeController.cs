using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Mail_WebArchiveView.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Message = "City of Gainesville Commissioner Email";

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Browse and Search City of Gainesville Commissioner Email.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Reach out to the Clerk of the Commission.";

            return View();
        }
    }
}
