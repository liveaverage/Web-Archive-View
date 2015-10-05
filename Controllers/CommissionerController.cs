using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Web.Configuration;

using System.IO;
using System.Xml;
using System.Xml.Serialization;

using log4net;
using log4net.Config;

using Mail_WebArchiveView.Models;
using PagedList;
using Microsoft.Exchange.WebServices.Data;


namespace Mail_WebArchiveView.Controllers

{
    public class CommissionerController : Controller
    {
        static readonly ILog logger = LogManager.GetLogger(typeof(CommissionerController));

        public ActionResult Index(int? page)
        {

            XmlSerializer xmlser = new XmlSerializer(typeof(xCommissioners));
            TextReader srdr = new StreamReader(Server.MapPath(Parameters.CommissionersXml));
            object obj = xmlser.Deserialize(srdr);
            xCommissioners Comms = (xCommissioners)obj;
            srdr.Close();

            return View(Comms.Commissioners);
            
        }

        public ActionResult Index2(int? page)
        {

            int pageNumber = (page ?? 1);
            List<Contact> CComms = GetComms();
            //return View(CComms.ToPagedList<Contact>(pageNumber, pageSize));
            return View(CComms);
        }

        public static List<Contact> GetComms()
        {
            ExchangeService service = Connection.ConnectEWS();
            ExpandGroupResults DG = service.ExpandGroup(ConfigurationManager.AppSettings["JournalDG"]);
            IEnumerable<EmailAddress> Members = DG.Members as IEnumerable<EmailAddress>;
            var Comms = new List<EmailAddress>();
            var CComms = new List<Contact>();

            //foreach (string s in Parameters.ArchiveExcludes)
            //{
            //    logger.Debug("Exclude: " + s);
            //}

            if (Parameters.ArchiveExcludes.Count() > 0)
            {
                
                foreach (EmailAddress i in Members)
                {
                    if (!(Parameters.ArchiveExcludes).Any(i.Address.Contains))
                    {
                        Comms.Add(i);

                        NameResolutionCollection coll = service.ResolveName(i.Name,
                                    ResolveNameSearchLocation.ContactsThenDirectory,
                                    true);

                        foreach (NameResolution res in coll)
                        {
                            Contact cont = res.Contact;
                            if (cont != null)
                            {
                                CComms.Add(cont);
                            }
                        }

                    }
                }
            }
            //Add explicit inclusions:
            if (Parameters.ArchiveInclude.Count > 0)
            {
                foreach(Contact i in Parameters.ArchiveInclude)
                {
                    if (i != null)
                    {
                        //logger.Debug("Include: " + i.DisplayName + " (" + i.JobTitle + ") " + i.EmailAddresses[EmailAddressKey.EmailAddress1].Address);
                        CComms.Add(i);
                    }
                }
            }

            return CComms;
        }

        public ActionResult CommissionerList()
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

            //var newlist = commList.ToList();

            XmlSerializer xmlser = new XmlSerializer(typeof(xCommissioners));
            TextReader srdr = new StreamReader(Server.MapPath(Parameters.CommissionersXml));
            object obj = xmlser.Deserialize(srdr);
            xCommissioners Comms = (xCommissioners)obj;
            srdr.Close();

            return PartialView("_CommissionerList", Comms.Commissioners);

            //return PartialView("_CommissionerList", newlist);
        }
    }


}
