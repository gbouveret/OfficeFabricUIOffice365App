using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeFabricUIAppForO365.Models
{
    public class EmailMessage
    {
        public DateTime Received { get; set; }
        public string Subject { get; set; }
        public string FromEmail { get; set; }
        public string FromName { get; set; }
        public bool HasAttachments { get; set; }
        public string Importance { get; set; }
        public bool IsRead { get; set; }
        public string Preview { get; set; }
    }
}