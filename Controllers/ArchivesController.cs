using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using System.Text;
using System.Diagnostics;

using log4net;
using log4net.Config;

using PagedList;
using EWS = Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using Mail_WebArchiveView.Models;

namespace Mail_WebArchiveView.Controllers
{
    public class ArchivesController : Controller
    {
        static readonly ILog logger = LogManager.GetLogger(typeof(ArchivesController));

        public ActionResult Index(int? page, string email, string name)
        {
            ExchangeService service = Connection.ConnectEWS();

            FolderView fv = new FolderView(50);
            SearchFilter folderSearch = new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, email);
            var findFoldersResults = service.FindFolders(WellKnownFolderName.Inbox, folderSearch, fv);
            Folder targetFolder = new Folder(service);
            
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

        public ActionResult Browse(string id, string name)
        {
            ViewBag.name = name;
            ViewBag.id = id;

            return View();
        }

        [HttpPost]
        public ActionResult Result(Search sc)
        {
            Session["search"] = sc;

            return View();
        }

        public ActionResult AjaxSearch(jQueryDataTableParamModel param)
        {
            ExchangeService service = Connection.ConnectEWS();
            FindItemsResults<Item> fMsgs;
            FindItemsResults<Item> aMsgs;
            List<EmailMessage> dMsgs = new List<EmailMessage>();

            Folder gsf = GetGlobalFolder(service);

            Search sc = null;

            #region Debug for Search Object:
            //if (sc != null)
            //{
            //    logger.Debug("Hit Ajax Search with Search criteria -- " + sc.subject + sc.to + sc.from + sc.datefrom + sc.dateto);

            //    if (sc.comms != null && sc.comms.Count() > 0)
            //    {
            //        foreach (string s in sc.comms)
            //        {
            //            logger.Debug("Search Commissioner: " + s);
            //        }
            //    }
            //}
            //else
            //{
            //    logger.Debug("AjaxSearch detected NULL SearchSession");
            //}
            #endregion

            string dateFilter;
            string fromFilter;
            string subjectFilter;
            string email;
            int totaldisplay;

            //Try getting the Datatables FILTER parameters or set to null. They're processed/tested later:
            try
            {
                dateFilter = Convert.ToString(Request["sSearch_0"]);
                fromFilter = Convert.ToString(Request["sSearch_1"]);
                subjectFilter = Convert.ToString(Request["sSearch_2"]);
                email = Convert.ToString(Request["email"]);
            }
            catch (Exception e)
            {
                dateFilter = null;
                fromFilter = null;
                subjectFilter = null;
                email = null;
                logger.Debug(e.Message);
            }


            if (!string.IsNullOrWhiteSpace(email))
            {
                //logger.Debug("Received email in querystring: " + email);
            }

            //logger.Debug("sSearch Params -- " + dateFilter + fromFilter + subjectFilter);

            if (sc != null)
            {
                //logger.Debug("Search form criteria -- " + sc.dateto + sc.datefrom);
              
                if ((!string.IsNullOrWhiteSpace(dateFilter) && !dateFilter.Equals("~")) || !string.IsNullOrWhiteSpace(fromFilter) || !string.IsNullOrWhiteSpace(subjectFilter))
                {
                    fMsgs = GetCMessages(service, gsf, sc.comms, param.iDisplayStart, param.iDisplayLength, 
                        sc.datefrom, sc.dateto, sc.to, sc.from, sc.subject, sc.bodytext, 
                        dateFilter, fromFilter, subjectFilter);
                    foreach (EmailMessage i in fMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = fMsgs.TotalCount;
                }
                else
                {
                    fMsgs = GetCMessages(service, gsf, sc.comms, param.iDisplayStart, param.iDisplayLength, 
                        sc.datefrom, sc.dateto, sc.to, sc.from, sc.subject, sc.bodytext, 
                        dateFilter, fromFilter, subjectFilter);
                    foreach (EmailMessage i in fMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = fMsgs.TotalCount;
                }

            }
            else if (!string.IsNullOrWhiteSpace(email))
            {
                logger.Debug("No session search detected. Browsing for: " + email);
                gsf = GetSourceFolder(service, email);

                if ((!string.IsNullOrWhiteSpace(dateFilter) && !dateFilter.Equals("~")) || !string.IsNullOrWhiteSpace(fromFilter) || !string.IsNullOrWhiteSpace(subjectFilter))
                {
                    string[] caemail = new string[] { email };

                    aMsgs = GetCMessages(service, gsf, caemail, param.iDisplayStart, param.iDisplayLength, null, null, null, null, null, null, dateFilter, fromFilter, subjectFilter);
                    foreach (EmailMessage i in aMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = aMsgs.TotalCount;
                }
                else
                {
                    logger.Debug("No filter params detected. Browsing for: " + email);

                    aMsgs = GetCMessages(service, gsf, param.iDisplayStart, param.iDisplayLength);

                    

                    foreach (EmailMessage i in aMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = aMsgs.TotalCount;
                    logger.Debug("Found messages for " + email + " : " + aMsgs.TotalCount + ". dMsgs count: " + dMsgs.Count);
                }
            }
            else
            {
                if ((!string.IsNullOrWhiteSpace(dateFilter) && !dateFilter.Equals("~")) || !string.IsNullOrWhiteSpace(fromFilter) || !string.IsNullOrWhiteSpace(subjectFilter))
                {
                    string[] c = {};
                    string s = "";
                    fMsgs = GetCMessages(service, gsf, c, param.iDisplayStart, param.iDisplayLength,
                        s, s, s, s, s, s,
                        dateFilter, fromFilter, subjectFilter);
                    foreach (EmailMessage i in fMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = fMsgs.TotalCount;
                }
                else
                {
                    aMsgs = GetCMessages(service, gsf, param.iDisplayStart, param.iDisplayLength);
                    foreach (EmailMessage i in aMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = aMsgs.TotalCount;
                }
            }

            
            var result = dMsgs.Select(x => new[] { x.DateTimeReceived.ToString(), x.From.Name, x.Subject, HttpUtility.UrlEncode(x.Id.ToString()), HttpUtility.UrlEncode(x.Id.ToString()) });
            
            return Json(new
            {
                sEcho = param.sEcho,
                iTotalRecords = gsf.TotalCount,
                iTotalDisplayRecords = totaldisplay,
                aaData = result 
            },
            JsonRequestBehavior.AllowGet);
        }

        public ActionResult AjaxSearchResults(jQueryDataTableParamModel param)
        {
            ExchangeService service = Connection.ConnectEWS();
            FindItemsResults<Item> fMsgs;
            FindItemsResults<Item> aMsgs;
            List<EmailMessage> dMsgs = new List<EmailMessage>();

            Folder gsf = GetGlobalFolder(service);

            Search sc = Session["search"] as Search;

            //Search sc = TempData["search"] as Search;

            if (sc != null)
            {
                logger.Debug("Hit Ajax Search with Search criteria -- " + sc.subject + sc.to + sc.from + sc.datefrom + sc.dateto);

                if (sc.comms != null && sc.comms.Count() > 0)
                {
                    foreach (string s in sc.comms)
                    {
                        logger.Debug("Search Commissioner: " + s);
                    }
                }
            }
            else
            {
                logger.Debug("AjaxSearch detected NULL SearchSession");
            }

            string dateFilter;
            string fromFilter;
            string subjectFilter;
            string email;
            int totaldisplay;

            //Try getting the Datatables FILTER parameters or set to null. They're processed/tested later:
            try
            {
                dateFilter = Convert.ToString(Request["sSearch_0"]);
                fromFilter = Convert.ToString(Request["sSearch_1"]);
                subjectFilter = Convert.ToString(Request["sSearch_2"]);
                email = Convert.ToString(Request["email"]);

                //sc = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<Search>(Request["query"]);
            }
            catch (Exception e)
            {
                dateFilter = null;
                fromFilter = null;
                subjectFilter = null;
                email = null;
                logger.Debug(e.Message);
            }


            if (!string.IsNullOrWhiteSpace(email))
            {
                logger.Debug("Received email in querystring: " + email);
            }

            logger.Debug("sSearch Params -- " + dateFilter + fromFilter + subjectFilter);

            if (sc != null)
            {
                logger.Debug("Search form criteria -- " + sc.dateto + sc.datefrom);

                if ((!string.IsNullOrWhiteSpace(dateFilter) && !dateFilter.Equals("~")) || !string.IsNullOrWhiteSpace(fromFilter) || !string.IsNullOrWhiteSpace(subjectFilter))
                {
                    fMsgs = GetCMessages(service, gsf, sc.comms, param.iDisplayStart, param.iDisplayLength,
                        sc.datefrom, sc.dateto, sc.to, sc.from, sc.subject, sc.bodytext,
                        dateFilter, fromFilter, subjectFilter);
                    foreach (EmailMessage i in fMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = fMsgs.TotalCount;
                }
                else
                {
                    fMsgs = GetCMessages(service, gsf, sc.comms, param.iDisplayStart, param.iDisplayLength,
                        sc.datefrom, sc.dateto, sc.to, sc.from, sc.subject, sc.bodytext,
                        dateFilter, fromFilter, subjectFilter);
                    foreach (EmailMessage i in fMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = fMsgs.TotalCount;
                }

            }
            else if (!string.IsNullOrWhiteSpace(email))
            {
                logger.Debug("No session search detected. Browsing for: " + email);
                gsf = GetSourceFolder(service, email);

                if ((!string.IsNullOrWhiteSpace(dateFilter) && !dateFilter.Equals("~")) || !string.IsNullOrWhiteSpace(fromFilter) || !string.IsNullOrWhiteSpace(subjectFilter))
                {
                    string[] caemail = new string[] { email };

                    aMsgs = GetCMessages(service, gsf, caemail, param.iDisplayStart, param.iDisplayLength, null, null, null, null, null, null, dateFilter, fromFilter, subjectFilter);
                    foreach (EmailMessage i in aMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = aMsgs.TotalCount;
                }
                else
                {
                    logger.Debug("No filter params detected. Browsing for: " + email);

                    aMsgs = GetCMessages(service, gsf, param.iDisplayStart, param.iDisplayLength);



                    foreach (EmailMessage i in aMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = aMsgs.TotalCount;
                    logger.Debug("Found messages for " + email + " : " + aMsgs.TotalCount + ". dMsgs count: " + dMsgs.Count);
                }
            }
            else
            {
                if ((!string.IsNullOrWhiteSpace(dateFilter) && !dateFilter.Equals("~")) || !string.IsNullOrWhiteSpace(fromFilter) || !string.IsNullOrWhiteSpace(subjectFilter))
                {
                    string[] c = { };
                    string s = "";
                    fMsgs = GetCMessages(service, gsf, c, param.iDisplayStart, param.iDisplayLength,
                        s, s, s, s, s, s,
                        dateFilter, fromFilter, subjectFilter);
                    foreach (EmailMessage i in fMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = fMsgs.TotalCount;
                }
                else
                {
                    aMsgs = GetCMessages(service, gsf, param.iDisplayStart, param.iDisplayLength);
                    foreach (EmailMessage i in aMsgs)
                    {
                        dMsgs.Add(i);
                    }
                    totaldisplay = aMsgs.TotalCount;
                }
            }


            var result = dMsgs.Select(x => new[] { x.DateTimeReceived.ToString(), x.From.Name, x.Subject, HttpUtility.UrlEncode(x.Id.ToString()), HttpUtility.UrlEncode(x.Id.ToString()) });

            return Json(new
            {
                sEcho = param.sEcho,
                iTotalRecords = gsf.TotalCount,
                iTotalDisplayRecords = totaldisplay,
                aaData = result
            },
            JsonRequestBehavior.AllowGet);
        }


        public ActionResult AjaxArchives (jQueryDataTableParamModel param)
        {
            logger.Debug("Parameters (iDisplayStart, iDisplayLength, sSearch, sEcho): " + param.iDisplayStart + "; " + param.iDisplayLength + "; " + param.sSearch + "; " + param.sEcho);
            ExchangeService service = Connection.ConnectEWS();
            FindItemsResults<Item> fMsgs;
            FindItemsResults<Item> aMsgs;
            Folder gsf;
            List<EmailMessage> dMsgs = new List<EmailMessage>();
            

            //Get search parameters:
            string dateFilter;
            string fromFilter;
            string subjectFilter;
            string email;

            try
            {
                 dateFilter = Convert.ToString(Request["sSearch_0"]);
                 fromFilter = Convert.ToString(Request["sSearch_1"]);
                 subjectFilter = Convert.ToString(Request["sSearch_2"]);
                 email = Convert.ToString(Request["email"]);
            }
            catch (Exception e)
            {
                dateFilter = null;
                fromFilter = null;
                subjectFilter = null;
                email = null;
                logger.Debug(e.Message);
            }

            if (!string.IsNullOrWhiteSpace(email))
            {
                logger.Debug("Received AjaxArchives Browse email: " + email);
                gsf = GetSourceFolder(service, email);
            }
            else
            {
                gsf = GetGlobalFolder(service);
            }
            
            int totaldisplay = gsf.TotalCount;

            logger.Debug("Search Parameters (iDisplayStart, iDisplayLenght, sSearch, sEcho): " + param.iDisplayStart + "; " + param.iDisplayLength + "; " + dateFilter + fromFilter + subjectFilter + "; " + param.sEcho);

            //List<EmailMessage> aMsgs = ArchivesController.GetMessages(param.iDisplayStart, param.iDisplayLength);

            if ((!string.IsNullOrWhiteSpace(dateFilter) && !dateFilter.Equals("~")) || !string.IsNullOrWhiteSpace(fromFilter) || !string.IsNullOrWhiteSpace(subjectFilter))
            {
                fMsgs = ArchivesController.GetAllMessages(service, gsf, param.iDisplayStart, param.iDisplayLength, dateFilter, fromFilter, subjectFilter);
                foreach (EmailMessage i in fMsgs)
                {
                    dMsgs.Add(i);
                }
                totaldisplay = fMsgs.TotalCount;
            }
            else
            {
                aMsgs = ArchivesController.GetAllMessages(service, gsf, param.iDisplayStart, param.iDisplayLength);
                foreach (EmailMessage i in aMsgs)
                {
                    dMsgs.Add(i);
                }
                totaldisplay = aMsgs.TotalCount;
            }


            var result = dMsgs.Select(x => new[] { x.DateTimeReceived.ToString(), x.From.Name, x.Subject, HttpUtility.UrlEncode(x.Id.ToString()) });

            return Json(new
            {
                sEcho = param.sEcho,
                iTotalRecords = gsf.TotalCount,
                iTotalDisplayRecords = totaldisplay,
                aaData = result
            },
        JsonRequestBehavior.AllowGet);
        }

        public ActionResult ajaxComm (jQueryDataTableParamModel param, string email)
        {
            logger.Debug("Parameters (iDisplayStart, iDisplayLength, sSearch, sEcho): " + param.iDisplayStart + "; " + param.iDisplayLength + "; " + param.sSearch + "; " + param.sEcho);
            ExchangeService service = Connection.ConnectEWS();
            Folder gsf = GetSourceFolder(service, email);
            FindItemsResults<Item> fMsgs;
            FindItemsResults<Item> aMsgs;
            List<EmailMessage> dMsgs = new List<EmailMessage>();
            int totaldisplay = gsf.TotalCount;

            //Get search parameters:
            string dateFilter;
            string fromFilter;
            string subjectFilter;

            try
            {
                dateFilter = Convert.ToString(Request["sSearch_0"]);
                fromFilter = Convert.ToString(Request["sSearch_1"]);
                subjectFilter = Convert.ToString(Request["sSearch_2"]);
            }
            catch (Exception e)
            {
                dateFilter = null;
                fromFilter = null;
                subjectFilter = null;
                logger.Debug(e.Message);
            }



            logger.Debug("Search Parameters (iDisplayStart, iDisplayLenght, sSearch, sEcho): " + param.iDisplayStart + "; " + param.iDisplayLength + "; " + dateFilter + fromFilter + subjectFilter + "; " + param.sEcho);

            //List<EmailMessage> aMsgs = ArchivesController.GetMessages(param.iDisplayStart, param.iDisplayLength);

            if ((!string.IsNullOrWhiteSpace(dateFilter) && !dateFilter.Equals("~")) || !string.IsNullOrWhiteSpace(fromFilter) || !string.IsNullOrWhiteSpace(subjectFilter))
            {
                fMsgs = ArchivesController.GetAllMessages(service, gsf, param.iDisplayStart, param.iDisplayLength, dateFilter, fromFilter, subjectFilter);
                foreach (EmailMessage i in fMsgs)
                {
                    dMsgs.Add(i);
                }
                totaldisplay = fMsgs.TotalCount;
            }
            else
            {
                aMsgs = ArchivesController.GetAllMessages(service, gsf, param.iDisplayStart, param.iDisplayLength);
                foreach (EmailMessage i in aMsgs)
                {
                    dMsgs.Add(i);
                }
                totaldisplay = aMsgs.TotalCount;
            }


            var result = dMsgs.Select(x => new[] { x.DateTimeReceived.ToString(), x.From.Name, x.Subject, HttpUtility.HtmlEncode(x.Id) });

            return Json(new
            {
                sEcho = param.sEcho,
                iTotalRecords = gsf.TotalCount,
                iTotalDisplayRecords = totaldisplay,
                aaData = result
            },
        JsonRequestBehavior.AllowGet);
        }

        public static Folder GetGlobalFolder(ExchangeService service)
        {

            FolderView fv = new FolderView(50);
            fv.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.TotalCount);
            SearchFilter folderSearch = new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, Parameters.GlobalSearchFolder);
            var findFoldersResults = service.FindFolders(WellKnownFolderName.SearchFolders, folderSearch, fv);
            Folder targetFolder = new Folder(service);

            foreach (Folder folder in findFoldersResults)
            {
                if (folder.DisplayName.ToLower() == Parameters.GlobalSearchFolder.ToLower())
                {
                    targetFolder = folder;
                }
            }

            return targetFolder;
        }

        private static Folder GetSourceFolder(ExchangeService service, string email)
        {
            // Use the following search filter to get all mail in the Inbox with the word "extended" in the subject line.
            SearchFilter searchCriteria = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                //Search for Folder DisplayName that matches mailbox email address:
                new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, email.ToLower()));
            FolderView fv = new FolderView(100);
            fv.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.TotalCount);

            // Find the search folder named "MailModder".
            FindFoldersResults findResults = service.FindFolders(
                WellKnownFolderName.Inbox,
                searchCriteria,
                fv);

            //Return root of inbox by default:
            //Folder returnFolder = Folder.Bind(service, WellKnownFolderName.Inbox);

            foreach (Folder searchFolder in findResults)
            {
                if (searchFolder.DisplayName.Contains(email.ToLower()))
                {
                    logger.Debug("Found source folder for email: " + email + ". Using folder: " + searchFolder.DisplayName + searchFolder.TotalCount);
                    return searchFolder;
                }
            }

            logger.Debug("Error getting commissioner source folder; returning global instead for: " + email);
            return GetGlobalFolder(service);
        }

        public static FindItemsResults<Item> GetAllMessages(ExchangeService service, Folder targetFolder, int start, int length)
        {

            //Create empty list for all mailbox messages:
            var listing = new List<EmailMessage>();

            //Create ItemView with correct pagesize and offset:
            ItemView view = new ItemView(length, start, OffsetBasePoint.Beginning);

            view.PropertySet = new PropertySet(EmailMessageSchema.Id,
                EmailMessageSchema.Subject,
                EmailMessageSchema.DateTimeReceived,
                EmailMessageSchema.From
                );

            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            FindItemsResults<Item> findResults = service.FindItems(targetFolder.Id, view);

            //bool MoreItems = true;

            //while(MoreItems)
            //{
            //foreach (EmailMessage it in findResults.Items)
            //{
            //    listing.Add(it);
            //}
            //}

            //return View(listing.ToPagedList<EmailMessage>(pageNumber, pageSize));

            return findResults;

        }

        public static FindItemsResults<Item> GetAllMessages(ExchangeService service, Folder targetFolder, int start, int length, string dateFilter, string fromFilter, string subjectFilter)
        {

            //Create empty list for all mailbox messages:
            var listing = new List<EmailMessage>();

            //Create ItemView with correct pagesize and offset:
            ItemView view = new ItemView(length, start, OffsetBasePoint.Beginning);

            view.PropertySet = new PropertySet(EmailMessageSchema.Id,
                EmailMessageSchema.Subject,
                EmailMessageSchema.DateTimeReceived,
                EmailMessageSchema.From
                );

            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
            List<SearchFilter> searchANDFilterCollection = new List<SearchFilter>();
            FindItemsResults<Item> findResults;

            if (!string.IsNullOrWhiteSpace(dateFilter))
            {
                string[] dates = dateFilter.Split('~');

                if (!string.IsNullOrWhiteSpace(dates[0]))
                {
                    searchANDFilterCollection.Add(new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, dates[0]));
                }

                if (!string.IsNullOrWhiteSpace(dates[1]))
                {
                    searchANDFilterCollection.Add(new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeReceived, dates[1]));
                }
            }

            if (!string.IsNullOrWhiteSpace(fromFilter))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.From, fromFilter));
            }

            if (!string.IsNullOrWhiteSpace(subjectFilter))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subjectFilter));
            }

            if (searchANDFilterCollection.Count > 0)
            {
                SearchFilter searchANDFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchANDFilterCollection.ToArray());
                findResults = service.FindItems(targetFolder.Id, searchANDFilter, view);
            }
            else
            {
                findResults = service.FindItems(targetFolder.Id, view);
            }

            logger.Debug("FindResults Count = " + findResults.TotalCount);

            //bool MoreItems = true;

            //while(MoreItems)
            //{
                //foreach (EmailMessage it in findResults.Items)
                //{
                //    listing.Add(it);
                //}
            //}

            //return View(listing.ToPagedList<EmailMessage>(pageNumber, pageSize));

            return findResults;

        }

        public static FindItemsResults<Item> GetCMessages(ExchangeService service, Folder targetFolder, int start, int length)
        {
            //Create empty list for all mailbox messages:
            var listing = new List<EmailMessage>();

            //Create ItemView with correct pagesize and offset:
            ItemView view = new ItemView(length, start, OffsetBasePoint.Beginning);

            view.PropertySet = new PropertySet(EmailMessageSchema.Id,
                EmailMessageSchema.Subject,
                EmailMessageSchema.DateTimeReceived,
                EmailMessageSchema.From
                );

            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            FindItemsResults<Item> findResults = service.FindItems(targetFolder.Id, view);

            //bool MoreItems = true;

            //while(MoreItems)
            //{
            //foreach (EmailMessage it in findResults.Items)
            //{
            //    listing.Add(it);
            //}
            //}

            //return View(listing.ToPagedList<EmailMessage>(pageNumber, pageSize));

            return findResults;

        }

        public static FindItemsResults<Item> GetCMessages(ExchangeService service, Folder targetFolder, string[] email, int start, int length, string datefromS, string datetoS, string toS, string fromS, string subjectS, string bodyS)
        {
            //Create empty list for all mailbox messages:
            var listing = new List<EmailMessage>();

            Folder gsf = GetGlobalFolder(service);
            FindItemsResults<Item> findResults = null;

            //Create ItemView with correct pagesize and offset:
            ItemView view = new ItemView(length, start, OffsetBasePoint.Beginning);

            view.PropertySet = new PropertySet(EmailMessageSchema.Id,
                EmailMessageSchema.Subject,
                EmailMessageSchema.DateTimeReceived,
                EmailMessageSchema.From
                );
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            //Define the new PidTagParentDisplay property to use for filtering:
            ExtendedPropertyDefinition def = new ExtendedPropertyDefinition(0x0E05, MapiPropertyType.String);

            List<SearchFilter> searchANDFilterCollection = new List<SearchFilter>();
            List<SearchFilter> searchORcommCollection = new List<SearchFilter>();
            List<SearchFilter> searchCompCollection = new List<SearchFilter>();

            //Add "OR" commissioner search criteria:
            if (email != null && email.Count() > 0)
            {
                foreach (var target in email)
                {
                    string t = target;
                    if (target.Contains('@'))
                    {
                        t = target.Split('@')[0];
                    }
                    
                    searchORcommCollection.Add(new SearchFilter.ContainsSubstring(def, t.ToLower()));

                    logger.Debug("Added mailbox to searchOR collection: " + t);
                }
            }

            if (searchORcommCollection.Count > 0)
            {
                SearchFilter searchOrFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchORcommCollection.ToArray());
                searchCompCollection.Add(searchOrFilter);
            }


            //Add all other search criteria

            if (!string.IsNullOrWhiteSpace(datefromS))
            {
                DateTime dFrom = DateTime.Parse(datefromS + " 12:00AM");
                logger.Debug("datefromS -- " + dFrom.ToString());

                searchANDFilterCollection.Add(new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, dFrom));
            }

            if (!string.IsNullOrWhiteSpace(datetoS))
            {
                DateTime dTo = DateTime.Parse(datetoS + " 11:59PM");
                logger.Debug("datetoS -- " + dTo.ToString());
                searchANDFilterCollection.Add(new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeReceived, dTo));
            }

            if (!string.IsNullOrWhiteSpace(fromS))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.From, fromS));
            }

            if (!string.IsNullOrWhiteSpace(toS))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.DisplayTo, toS));
            }

            if (!string.IsNullOrWhiteSpace(subjectS))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subjectS));
            }

            if (!string.IsNullOrWhiteSpace(bodyS))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Body, bodyS));
            }

            if (searchANDFilterCollection.Count > 0)
            {
                SearchFilter searchANDFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchANDFilterCollection.ToArray());
                searchCompCollection.Add(searchANDFilter);
            }

            //Evaluate filters and execute find results:

            if (searchORcommCollection.Count > 0 || searchANDFilterCollection.Count > 0)
            {
                logger.Debug("FindResults execution with comp search collection");
                SearchFilter searchComp = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchCompCollection.ToArray());
                findResults = service.FindItems(targetFolder.Id, searchComp, view);
            }
            else
            {
                findResults = service.FindItems(targetFolder.Id, view);
            }

            logger.Debug("FindResults Count = " + findResults.TotalCount);

            
            return findResults;

        }

        public static FindItemsResults<Item> GetCMessages(ExchangeService service, Folder targetFolder, string[] email, int start, int length, string datefromS, string datetoS, string toS, string fromS, string subjectS, string bodyS, string dateFilter, string fromFilter, string subjectFilter)
        {
            //Create empty list for all mailbox messages:
            var listing = new List<EmailMessage>();

            Folder gsf = gsf = GetGlobalFolder(service);
            FindItemsResults<Item> findResults = null;

            //Create ItemView with correct pagesize and offset:
            ItemView view = new ItemView(length, start, OffsetBasePoint.Beginning);

            view.PropertySet = new PropertySet(EmailMessageSchema.Id,
                EmailMessageSchema.Subject,
                EmailMessageSchema.DateTimeReceived,
                EmailMessageSchema.From
                );
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            //Define the new PidTagParentDisplay property to use for filtering:
            ExtendedPropertyDefinition def = new ExtendedPropertyDefinition(0x0E05, MapiPropertyType.String);

            List<SearchFilter> searchANDFilterCollection = new List<SearchFilter>();
            List<SearchFilter> searchORcommCollection = new List<SearchFilter>();
            List<SearchFilter> searchCompCollection = new List<SearchFilter>();

            //Add "OR" commissioner search criteria:
            if (email != null && email.Count() > 0)
            {
                foreach (var target in email)
                {
                    string t = target;
                    if (target.Contains('@'))
                    {
                        t = target.Split('@')[0];
                    }

                    searchORcommCollection.Add(new SearchFilter.ContainsSubstring(def, t.ToLower()));
                    logger.Debug("Added mailbox to searchOR collection: " + target);
                }
            }

            if (searchORcommCollection.Count > 0)
            {
                SearchFilter searchOrFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchORcommCollection.ToArray());
                searchCompCollection.Add(searchOrFilter);
            }


            //Populate fields from the SEARCH form (not the filters, just the search):

            if (!string.IsNullOrWhiteSpace(datefromS))
            {
                DateTime dFrom = DateTime.Parse(datefromS + " 12:00AM");
                logger.Debug("datefromS -- " + dFrom.ToString());

                searchANDFilterCollection.Add(new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, dFrom));
            }

            if (!string.IsNullOrWhiteSpace(datetoS))
            {
                DateTime dTo = DateTime.Parse(datetoS + " 11:59PM");
                logger.Debug("datetoS -- " + dTo.ToString());
                searchANDFilterCollection.Add(new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeReceived, dTo));
            }

            if (!string.IsNullOrWhiteSpace(fromS))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.From, fromS));
            }

            if (!string.IsNullOrWhiteSpace(toS))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.DisplayTo, toS));
            }

            if (!string.IsNullOrWhiteSpace(subjectS))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subjectS));
            }

            if (!string.IsNullOrWhiteSpace(bodyS))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Body, bodyS));
            }

            //Populate fields from Datatables FILTER form (this supplements the SEARCH form):

            if (!string.IsNullOrWhiteSpace(dateFilter))
            {
                string[] dates = dateFilter.Split('~');

                if (!string.IsNullOrWhiteSpace(dates[0]))
                {
                    DateTime dfFrom = DateTime.Parse(dates[0] + " 12:00AM");
                    logger.Debug("dfFrom -- " + dfFrom.ToString());
                    searchANDFilterCollection.Add(new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, dfFrom));
                }

                if (!string.IsNullOrWhiteSpace(dates[1]))
                {
                    DateTime dfTo = DateTime.Parse(dates[1] + " 11:59PM");
                    logger.Debug("dfTo -- " + dfTo.ToString());
                    searchANDFilterCollection.Add(new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeReceived, dfTo));
                }
            }

            if (!string.IsNullOrWhiteSpace(fromFilter))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.From, fromFilter));
            }

            if (!string.IsNullOrWhiteSpace(subjectFilter))
            {
                searchANDFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subjectFilter));
            }

            //Assemble the SearchFilter Collection

            if (searchANDFilterCollection.Count > 0)
            {
                SearchFilter searchANDFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchANDFilterCollection.ToArray());
                searchCompCollection.Add(searchANDFilter);
            }

            //Evaluate filters and execute find results:

            if (searchORcommCollection.Count > 0 || searchANDFilterCollection.Count > 0)
            {
                SearchFilter searchComp = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchCompCollection.ToArray());
                findResults = service.FindItems(targetFolder.Id, searchComp, view);
            }
            else
            {
                findResults = service.FindItems(targetFolder.Id, view);
            }

            logger.Debug("FindResults Count = " + findResults.TotalCount);


            return findResults;

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
