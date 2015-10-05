using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Exchange.WebServices.Data;
using System.Xml;
using System.Xml.Serialization;


namespace Mail_WebArchiveView.Models
{
    public class Commissioner
    {
        public string[] SelectedIds { get; set; }
        public List<Contact> Contacts { get; set; }
    }
    [XmlRoot("Commissioner")]
    public class xCommissioner
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string Mail { get; set; }
        public string CurrentCommissioner { get; set; }
    }

    [XmlRoot("ArrayOfCommissioner")]
    public class xCommissioners
    {
        [XmlElement("Commissioner")]
        public List<xCommissioner> Commissioners { get; set; }
    }
}   