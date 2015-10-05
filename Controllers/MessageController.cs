using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net;
using ICSharpCode.SharpZipLib.Zip;

using System.IO;
using System.IO.Compression;

using log4net;
using log4net.Config;

using Mail_WebArchiveView.Models;
using PagedList;
using Microsoft.Exchange.WebServices.Data;

namespace Mail_WebArchiveView.Controllers
{
    public class MessageController : Controller
    {
        //
        // GET: /Message/
        static readonly ILog logger = LogManager.GetLogger(typeof(CommissionerController));

        public ActionResult Index(string mid, string email, string name, bool preview)
        {
            ExchangeService service = Connection.ConnectEWS();
            EmailMessage message = EmailMessage.Bind(service, new ItemId(mid), new PropertySet(BasePropertySet.FirstClassProperties,
                EmailMessageSchema.Body,
                EmailMessageSchema.From,
                EmailMessageSchema.DisplayTo,
                EmailMessageSchema.DisplayCc,
                EmailMessageSchema.BccRecipients,
                ItemSchema.Attachments));
            message.Load();

            string att = "";


            //Only Inventory attachments; don't load unless needed.
            foreach (Attachment attachment in message.Attachments)
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = attachment as FileAttachment;
                    att += (fileAttachment.Name + ", ");

                }
                else // Attachment is an item attachment.
                {
                    // Load attachment into memory and write out the subject.
                    ItemAttachment itemAttachment = attachment as ItemAttachment;
                    att += (itemAttachment.Name + ", ");
                }
            }

            //Define ViewBag.email and TO recipient string:
            ViewBag.email = email;
            ViewBag.att = att.TrimEnd(',');
            ViewBag.name = name;

            //Default ActionResult:
            ActionResult v = View(message);

            if (preview == false)
            {
                ViewBag.preview = false;
                v = View("Index", message);
            }
            else
            {
                ViewBag.preview = true;
                v = View("Index", "Msg" ,message);
            }

            return (v);
        }

        public ActionResult GetAttachment(string aid, string attname, string mid)
        {
            ExchangeService service = Connection.ConnectEWS();
            EmailMessage message = EmailMessage.Bind(service, new ItemId(mid), new PropertySet(ItemSchema.Attachments));
            message.Load();
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            byte[] ContentBytes = null;
            string filename = null;
            string contentType = null;

            foreach (Attachment att in message.Attachments)
            {
                if (aid == att.Id.ToString() && att is FileAttachment)
                {
                    FileAttachment fileAt = att as FileAttachment;
                    fileAt.Load(ms);
                    filename = fileAt.Name;
                    contentType = fileAt.ContentType;
                }
                else if (aid == att.Id.ToString() && att is ItemAttachment)
                {
                    ItemAttachment itemAttachment = att as ItemAttachment;
                    itemAttachment.Load(new PropertySet(EmailMessageSchema.MimeContent));
                    MimeContent mc = itemAttachment.Item.MimeContent;
                    ContentBytes = mc.Content;
                    filename = itemAttachment.Name + ".eml";
                    contentType = itemAttachment.ContentType;

                }
            }

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());

            // return the file
            return File(ms.ToArray() ?? ContentBytes, attname);
        }

        public ActionResult GetMessageDL(string mid)
        {
            
            byte[] ContentBytes = null;
            string filename = null;
            
            ExchangeService service = Connection.ConnectEWS();
            PropertySet ps = new PropertySet(ItemSchema.MimeContent, ItemSchema.DateTimeReceived);
            EmailMessage message = EmailMessage.Bind(service, new ItemId(mid), ps);
            message.Load(ps);
            MimeContent mc = message.MimeContent;
            ContentBytes = mc.Content;
            filename = string.Format("COG_{0:yyyyMMddhh_mmsstt}.eml", message.DateTimeReceived);

            //foreach (Attachment att in message.Attachments)
            //{
            //    if (aid == att.Id.ToString() && att is FileAttachment)
            //    {
            //        FileAttachment fileAt = att as FileAttachment;
            //        fileAt.Load(ms);
            //        filename = fileAt.Name;
            //        contentType = fileAt.ContentType;
            //    }
            //    else if (aid == att.Id.ToString() && att is ItemAttachment)
            //    {
            //        ItemAttachment itemAttachment = att as ItemAttachment;
            //        itemAttachment.Load(new PropertySet(EmailMessageSchema.MimeContent));
            //        MimeContent mc = itemAttachment.Item.MimeContent;
            //        ContentBytes = mc.Content;
            //        filename = itemAttachment.Name + ".eml";
            //        contentType = itemAttachment.ContentType;

            //    }
            //}

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());

            // return the file
            return File(ContentBytes, filename);
        }

        public byte[] GetMessageEntry(string mid)
        {
            byte[] ContentBytes = null;
            string filename = null;
            ExchangeService service = Connection.ConnectEWS();
            PropertySet ps = new PropertySet(ItemSchema.MimeContent, ItemSchema.DateTimeReceived);

            EmailMessage message = EmailMessage.Bind(service, new ItemId(mid), ps);
            message.Load(ps);
            MimeContent mc = message.MimeContent;
            ContentBytes = mc.Content;
            filename = string.Format("COG_{0:yyyyMMddhh_mmsstt}.eml", message.DateTimeReceived);

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = filename,
                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());

            // return the file
            return ContentBytes;
        }

        public ActionResult GetMessageZip(ICollection<string> msga)
        {
            if (msga == null || msga.Count() <= 0)
            {
                return new EmptyResult();
                //return RedirectToAction("Index", "Search");
            }


            ExchangeService service = Connection.ConnectEWS();
            PropertySet ps = new PropertySet(ItemSchema.MimeContent, ItemSchema.DateTimeReceived);
            string zfilename = (DateTime.Now.ToString() + ".zip");


            using (var compressedFileStream = new MemoryStream())
            {
                //Create an archive and store the stream in memory.
                using (var zipArchive = new ZipArchive(compressedFileStream, ZipArchiveMode.Update, false))
                {
                    foreach (var mid in msga)
                    {
                        byte[] ContentBytes = null;
                        string filename = null;

                        logger.Debug("Bulk Export for: " + mid);
                        logger.Debug("Decoded: " + WebUtility.UrlDecode(mid).ToString());

                        EmailMessage message = EmailMessage.Bind(service, new ItemId(WebUtility.UrlDecode(mid)), ps);
                        message.Load(ps);
                        MimeContent mc = message.MimeContent;
                        ContentBytes = mc.Content;
                        filename = string.Format("COG_{0:yyyyMMddhh_mmsstt}.eml", message.DateTimeReceived);

                        //Create a zip entry for each message
                        var zipEntry = zipArchive.CreateEntry(filename);

                        //Get the stream of the messages MIME content
                        using (var originalFileStream = new MemoryStream(ContentBytes))
                        {
                            using (var zipEntryStream = zipEntry.Open())
                            {
                                //Copy the message stream to the zip entry stream
                                originalFileStream.CopyTo(zipEntryStream);
                            }
                        }
                    }
                }
                var cd = new System.Net.Mime.ContentDisposition
                {
                    FileName = zfilename,
                    // always prompt the user for downloading, set to true if you want 
                    // the browser to try to show the file inline
                    Inline = false,
                };
                Response.AppendHeader("Content-Disposition", cd.ToString());


                return File(compressedFileStream.ToArray(), zfilename);
                //return new FileContentResult(compressedFileStream.ToArray(), "application/zip") { FileDownloadName = zfilename };
            }
        }

        // Compresses the supplied memory stream, naming it as zipEntryName, into a zip,
        // which is returned as a memory stream or a byte array.
        //
        //public ActionResult CreateToMemoryStream(string[] mids, MemoryStream memStreamIn, string zipEntryName)
        //{
        //    if (mids.Count() <= 0)
        //    {
        //        return View();
        //    }

        //    ExchangeService service = Connection.ConnectEWS();
        //    PropertySet ps = new PropertySet(ItemSchema.MimeContent, ItemSchema.DateTimeReceived);
        //    string zfilename = (DateTime.Now.ToString() + ".zip");

        //    MemoryStream outputMemStream = new MemoryStream();
        //    ZipOutputStream zipStream = new ZipOutputStream(outputMemStream);

        //    zipStream.SetLevel(3); //0-9, 9 being the highest level of compression
            
            
        //    //Loop this:
        //    foreach (var mid in mids)
        //    {
        //        byte[] ContentBytes = null;
        //        string filename = null;

        //        EmailMessage message = EmailMessage.Bind(service, new ItemId(mid), ps);
        //        message.Load(ps);
        //        MimeContent mc = message.MimeContent;
        //        ContentBytes = mc.Content;
        //        filename = string.Format("COG_{0:yyyyMMddhh_mmsstt}.eml", message.DateTimeReceived);

        //        ZipEntry newEntry = new ZipEntry(filename);
        //        newEntry.DateTime = DateTime.Now;

        //        zipStream.PutNextEntry(newEntry);

        //        ICSharpCode.SharpZipLib.Core.StreamUtils.Copy(new MemoryStream(ContentBytes), zipStream, new byte[4096]);
        //        zipStream.CloseEntry();
        //    }
        //    //End Loop

        //    zipStream.IsStreamOwner = false;    // False stops the Close also Closing the underlying stream.
        //    zipStream.Close();          // Must finish the ZipOutputStream before using outputMemStream.

        //    outputMemStream.Position = 0;
        //    //return outputMemStream;

        //    var cd = new System.Net.Mime.ContentDisposition
        //    {
        //        FileName = zfilename,
        //        // always prompt the user for downloading, set to true if you want 
        //        // the browser to try to show the file inline
        //        Inline = false,
        //    };
        //    Response.AppendHeader("Content-Disposition", cd.ToString());
            
        //    return File(zipStream, zfilename);

        //    // Alternative outputs:
        //    // ToArray is the cleaner and easiest to use correctly with the penalty of duplicating allocated memory.
        //    byte[] byteArrayOut = outputMemStream.ToArray();

        //    // GetBuffer returns a raw buffer raw and so you need to account for the true length yourself.
        //    //byte[] byteArrayOut = outputMemStream.GetBuffer();
        //    long len = outputMemStream.Length;
        //}
    }
}
