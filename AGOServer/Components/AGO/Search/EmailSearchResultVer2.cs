using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class EmailSearchResultVer2
    {
        List<EmailSearchInfoVer2> emailInfos = new List<EmailSearchInfoVer2>();
        int listHead;
        int includeCount;
        int actualCount;
        string remarks;

        public List<EmailSearchInfoVer2> EmailInfos { get => emailInfos; set => emailInfos = value; }
        public int PageNumber { get; set; }
        public int PageSize { get; set; }
        public int TotalPage { get; set; }
        public int MaxCount { get; set; }
        public int TotalCount { get; set; }
        public bool IsMaxCountReached { get; set; }
        public string SortedBy { get; internal set; }
        public string SortDirection { get; internal set; }
        public int ListHead { get => listHead; set => listHead = value; }
        public int IncludeCount { get => includeCount; set => includeCount = value; }
        public int ActualCount { get => actualCount; set => actualCount = value; }
        public string Remarks { get => remarks; set => remarks = value; }
    }
}