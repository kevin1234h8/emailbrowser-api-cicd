using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer.Components
{
    public class ListPaginatedEmailsResult
    {
        public List<EmailInfo> EmailInfos { get; set; }
        public int NumberOfPages { get; set; }
        public int PageNumber { get; set; }
        public int TotalCount { get; set; }
        public int PageSize { get; internal set; }
        public string SortedBy { get; internal set; }
        public string SortDirection { get; internal set; }
    }
}