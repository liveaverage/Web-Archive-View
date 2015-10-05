using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net;
using System.Security.Cryptography;
using System.Text;

using log4net;
using log4net.Config;

using PagedList;
using EWS = Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using Mail_WebArchiveView.Models;

namespace Mail_WebArchiveView.Models
{
    public class Parameters : Controller
    {

        public static string[] ArchiveExcludes
        {
            get
            {
                string[] a = null;
                if (!string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings["ArchiveExcludes"]))
                {
                    a = (ConfigurationManager.AppSettings["ArchiveExcludes"]).Split(';');
                }

                return a;
            }
        }
        
        //public static List<EmailAddress> ArchiveIncludes
        //{
        //    get
        //    {
        //        List<EmailAddress> l = new List<EmailAddress>();
        //        if (!string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings["ArchiveIncludes"]))
        //        {
        //            string[] users = (ConfigurationManager.AppSettings["ArchiveIncludes"]).Split(';');

        //            if (users.Count() > 0)
        //            {
        //                foreach (var i in users)
        //                {
        //                    string[] usr = i.Split('/');
        //                    EmailAddress ne = new EmailAddress(usr[0], usr[1]);
        //                    ne.Id = usr[2];
        //                    l.Add(ne);
        //                    ne = null;
        //                }
        //            }
        //        }
        //        return (l);
        //    }
        //}

        public static List<Contact> ArchiveInclude
        {
            get
            {
                ExchangeService service = Connection.ConnectEWS();
                List<Contact> l = new List<Contact>();
                if (!string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings["ArchiveIncludes"]))
                {
                    
                    string[] users = (ConfigurationManager.AppSettings["ArchiveIncludes"]).Split(';');

                    if (users.Count() > 0)
                    {
                        foreach (var i in users)
                        {
                            if (!String.IsNullOrWhiteSpace(i))
                            {
                                string[] usr = i.Split('/');
                                Contact c = new Contact(service);
                                c.DisplayName = usr[0];
                                c.EmailAddresses[EmailAddressKey.EmailAddress1] = usr[1];
                                c.JobTitle = usr[2];
                                l.Add(c);
                                c = null;
                            }
                        }
                    }
                }
                return (l);
            }
            
        }

        public static string GlobalSearchFolder
        {
            get
            {
                string a = "MailModder";

                if (!string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings["ExGlobalSearch"]))
                {
                    a = (ConfigurationManager.AppSettings["ExGlobalSearch"]);
                }

                return a;
            }
        }

        public static string GlobalMailDomain
        {
            get
            {
                string md = null;

                if (!string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings["MailDomain"]))
                {
                    md = (ConfigurationManager.AppSettings["MailDomain"]);
                }

                return md;
            }

        }

        public static string CommissionersXml
        {
            get
            {
                string md = "~/Commissioners.xml";

                if (!string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings["CommissionersXml"]))
                {
                    md = (ConfigurationManager.AppSettings["CommissionersXml"]);
                }

                return md;
            }
        }
    }
}
