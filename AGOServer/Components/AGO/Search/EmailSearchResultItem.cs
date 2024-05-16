using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class EmailSearchResultItem
    {
        public string Summary { get; internal set; }
        public string FullLocation { get; internal set; }
        public string FileSize { get; internal set; }
        public string Name { get; internal set; }
    }
}