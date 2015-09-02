using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeFabricUIAppForO365.Models
{
    public class Persona
    {
        public string DisplayName { get; set; }
        public string Email { get; set; }      
        public string JobTitle { get; set; } 
        public string CompanyName { get; set; }
        public string OfficeLocation { get; set; }
        public string Phone { get; set; }
        public string Im { get; set; }

    }
}