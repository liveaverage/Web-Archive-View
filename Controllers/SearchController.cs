using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Exchange.WebServices.Data;
using MvcCheckBoxList;
using Mail_WebArchiveView.Models;

using log4net;
using log4net.Config;

using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace Mail_WebArchiveView.Controllers
{
    public class SearchController : Controller
    {
        //
        // GET: /Search/
        static readonly ILog logger = LogManager.GetLogger(typeof(ArchivesController));

        public ActionResult Index()
        {
            //List<Contact> l = CommissionerController.GetComms();
            //List<EmailAddress> commList = new List<EmailAddress>();

            //foreach (var item in l)
            //{
            //    List<string> eads = new List<string>();
                

            //    EmailAddress t = new EmailAddress();
            //    t.Name = item.DisplayName;
            //    t.Id = item.JobTitle;
            //    if (item.EmailAddresses != null)
            //    {

            //        if (item.EmailAddresses.Contains(EmailAddressKey.EmailAddress1))
            //        {
            //            eads.Add(item.EmailAddresses[EmailAddressKey.EmailAddress1].Address);
            //        }

            //        if (item.EmailAddresses.Contains(EmailAddressKey.EmailAddress2))
            //        {
            //            eads.Add(item.EmailAddresses[EmailAddressKey.EmailAddress2].Address);
            //        }

            //        if (item.EmailAddresses.Contains(EmailAddressKey.EmailAddress3))
            //        {
            //            eads.Add(item.EmailAddresses[EmailAddressKey.EmailAddress3].Address);
            //        }
            //    }

            //    foreach (string s in eads)
            //    {
            //        if (s.StartsWith("SMTP:", StringComparison.CurrentCulture))
            //        {
            //            t.Address = (s.Split(':')[1]);
            //        }

            //    }

            //    commList.Add(t);
            //}

            XmlSerializer xmlser = new XmlSerializer(typeof(xCommissioners));
            TextReader srdr = new StreamReader(Server.MapPath(Parameters.CommissionersXml));
            object obj = xmlser.Deserialize(srdr);
            xCommissioners Comms = (xCommissioners)obj;
            srdr.Close();

            Session["search"] = null;

            return View(Comms.Commissioners);
        }

        public ActionResult Index2()
        {
            List<Contact> l = CommissionerController.GetComms();
            return View(l);
        }

        public PartialViewResult CommissionerChecklist()
        {
            List<Contact> l = CommissionerController.GetComms();
            return PartialView(l);
        }

    }
}
