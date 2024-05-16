using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class EmailContentInfo
    {
        public string FileName { get; set; }
        public DateTime? CreationDate { get; set; }
        public DateTime? ModificationDate { get; set; }
        public long? FileSize { get; set; }
    }
}