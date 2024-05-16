using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class EmailSearchResult
    {
        List<EmailInfo> emailInfos = new List<EmailInfo>();
        int listHead;
        int includeCount;
        int actualCount;
        string remarks;
        string cache_id;

        public List<EmailInfo> EmailInfos { get => emailInfos; set => emailInfos = value; }
        public int ListHead { get => listHead; set => listHead = value; }
        public int IncludeCount { get => includeCount; set => includeCount = value; }
        public int ActualCount { get => actualCount; set => actualCount = value; }
        public string Remarks { get => remarks; set => remarks = value; }
        public string CacheID { get => cache_id; set => cache_id = value; }
    }
}