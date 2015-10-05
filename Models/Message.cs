using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace Mail_WebArchiveView.Models
{

    public class Message
    {
        public string ID { get; set; }
        public DateTime Received { get; set ; }
        public string FromAdd { get; set; }
        public string ToAdd { get; set; }
        public string Subject { get; set; }
    }

}