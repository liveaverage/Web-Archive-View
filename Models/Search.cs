using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Exchange.WebServices.Data;
namespace Mail_WebArchiveView.Models
{
    public class Search
    {
        public string[] comms { get; set;}
        public string datefrom { get; set;}
        public string dateto { get; set;}
        public string from { get; set;} 
        public string to { get; set;}
        public string subject { get; set;}
        public string bodytext { get; set;}
    }
}