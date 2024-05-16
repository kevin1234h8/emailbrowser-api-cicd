using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer.Components.Models.OpenText
{
    public class OTDSCurrentUserResponse
    {
        public bool isAdmin { get; set; }
        public Dictionary<string, object> user { get; set; }
        public bool isSysAdmin { get; set; }
    }
}