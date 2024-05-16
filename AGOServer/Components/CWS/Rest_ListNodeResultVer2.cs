using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class Rest_ListNodeResultVer2
    {
        public List<Rest_NodeVer2> Nodes { get; set; }
        public int PageNumber { get; set; }
        public int PageSize { get; set; }        
        public int TotalCount { get; set; }
        public string ActualSortedBy { get; set; }
        public int TotalPage { get; internal set; }
        public int RangeMin { get; internal set; }
        public int RangeMax { get; internal set; }
    }
}