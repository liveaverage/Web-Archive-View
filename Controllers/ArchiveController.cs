using System;
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

//Needed for JSON output to jQuery:
using System.Web.Script.Serialization;


namespace Mail_WebArchiveView.Controllers
{
    public class ArchiveController : Controller
    {

        public ActionResult Index(int? page, string email, string name)
        {
            ExchangeService service = Connection.ConnectEWS();

            FolderView fv = new FolderView(50);
            SearchFilter folderSearch = new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, email.ToLower());
            var findFoldersResults = service.FindFolders(WellKnownFolderName.Inbox, folderSearch, fv);
            Folder targetFolder = null;

            foreach (Folder folder in findFoldersResults)
            {
                if (folder.DisplayName.ToLower() == email.ToLower())
                {
                    targetFolder = folder;
                }
            }
            //Define ViewBag.email:
            ViewBag.email = email;
            ViewBag.emailnq = email.Split('@')[0];
            ViewBag.name = name;

            //Create empty list for all mailbox messages:
            var listing = new List<EmailMessage>();

            //Create ItemView with correct pagesize and offset:
            ItemView view = new ItemView(int.MaxValue, Connection.ExOffset, OffsetBasePoint.Beginning);

            view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties,
                EmailMessageSchema.Subject,
                EmailMessageSchema.DateTimeReceived,
                EmailMessageSchema.From
                );

            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            FindItemsResults<Item> findResults = service.FindItems(targetFolder.Id, view);

            //bool MoreItems = true;

            //while(MoreItems)
            //{
            foreach (EmailMessage it in findResults.Items)
            {
                listing.Add(it);
            }
            //}
            int pageSize = Connection.ExPageSize;
            int pageNumber = (page ?? 1);
            //return View(listing.ToPagedList<EmailMessage>(pageNumber, pageSize));
            return View(listing.ToList<EmailMessage>());
        }

        public string Welcome(string name, int numTime = 1)
        {
            //return "This is the Welcome action method...";
            return HttpUtility.HtmlEncode("Hello " + name + ", NumTimes is: " + numTime);
        }

        public ActionResult Messages()
        {
            RDOSession Session = new RDOSession();
            Session.Logon();
            RDOExchangeMailboxStore mbStore = (RDOExchangeMailboxStore)Session.Stores.DefaultStore;

            //Evaluate with Drafts folder:
            RDOFolder Inbox = (RDOFolder)Session.GetDefaultFolder(rdoDefaultFolders.olFolderDrafts);

            var listing = new List<Message>();

            foreach (RDOMail item in Inbox.Items)
            {
                var temp = new Message();
                temp.Subject = item.Subject;
                temp.Received = item.ReceivedTime;
                temp.ID = item.EntryID;
                listing.Add(temp);
                temp = null;
            }
            
            GC.Collect();

            return View(listing.ToList<Message>());
        }

        public ActionResult MessageEWS(EmailAddress email)
        {
            ExchangeService service = Connection.ConnectEWS();

            //Create empty list for all mailbox messages:
            var listing = new List<EmailMessage>();

            //Create ItemView with correct pagesize and offset:
            ItemView view = new ItemView(Connection.ExPageSize, Connection.ExOffset, OffsetBasePoint.Beginning);

            view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties,
                EmailMessageSchema.Subject,
                EmailMessageSchema.DateTimeReceived,
                EmailMessageSchema.From,
                EmailMessageSchema.ToRecipients);

            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            //string sf = "Body:\"Compensation\"";

            //Define the new PidTagParentDisplay property to use for filtering:
            ExtendedPropertyDefinition def = new ExtendedPropertyDefinition(0x0E05, MapiPropertyType.String);
            SearchFilter searchCriteria = new SearchFilter.IsEqualTo(def, email.Address);
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, searchCriteria, view);

            foreach (EmailMessage it in findResults.Items)
            {
                listing.Add(it);
            }

            return View(listing.ToList<EmailMessage>());
        }

        private static Folder GetSourceFolder(ExchangeService service, EmailAddress email)
        {
            log4net.Config.XmlConfigurator.Configure();
           
            // Use the following search filter to get all mail in the Inbox with the word "extended" in the subject line.
            SearchFilter searchCriteria = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                //Search for Folder DisplayName that matches mailbox email address:
                new SearchFilter.IsEqualTo(FolderSchema.DisplayName, email.Address));

            // Find the search folder named "MailModder".
            FindFoldersResults findResults = service.FindFolders(
                WellKnownFolderName.Inbox,
                searchCriteria,
                new FolderView(50));
            
            //Return root of inbox by default:
            Folder returnFolder = Folder.Bind(service, WellKnownFolderName.Inbox);
            foreach (Folder searchFolder in findResults.Folders)
            {
                if (searchFolder.DisplayName == email.Address)
                {
                    returnFolder = searchFolder;
                }
            }

            return returnFolder;
        }

        public static bool CertificateValidationCallBack(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

    }

}
