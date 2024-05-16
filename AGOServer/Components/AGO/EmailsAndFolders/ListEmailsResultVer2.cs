using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer.Components
{
    public class ListEmailsResultVer2
    {
        public List<EmailInfoVer2> EmailInfos { get; set; }
        public int PageNumber { get; set; }
        public int PageSize { get; set; }
        public int TotalPage { get; set; }
        public int MaxCount { get; set; }
        public int TotalCount { get; set; }
        public bool IsMaxCountReached { get; set; }
        public string SortedBy { get; internal set; }
        public string SortDirection { get; internal set; }
    }
}