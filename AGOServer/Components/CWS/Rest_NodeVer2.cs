using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class Rest_NodeVer2
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public long ParentId { get; internal set; }
        public long Size { get; internal set; }
        public int Type { get; internal set; }
        public string ConversationId { get; set; }
        public string EmailSubject { get; set; }
        public string EmailTo { get; set; }
        public string EmailCc { get; set; }
        public string EmailFrom { get; set; }
        public DateTime SentDate { get; set; }
        public DateTime ReceivedDate { get; set; }
        public int HasAttachments { get; set; }
    }
}