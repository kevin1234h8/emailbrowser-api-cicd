using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class ExtractedEmailInfo
    {
        public long CSID { get; internal set; }
        public string CSName { get; internal set; }
        public long CSVersionNum { get; internal set; }
        public DateTime CSModifyDate { get; internal set; }
        public long CSParentID { get; internal set; }
        public long CSCachingFolderNodeID { get; internal set; }
        public bool IsExtracted { get; internal set; }
        public bool IsExtractingNow { get; internal set; }
        public DateTime LastExtractedDate { get; internal set; }
    }
}