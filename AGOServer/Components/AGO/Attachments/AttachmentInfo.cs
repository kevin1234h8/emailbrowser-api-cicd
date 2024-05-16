using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AGOServer
{
    public class AttachmentInfo
    {
        public int No { get; set; }
        public long CSEmailID { get ; set; }
        public byte[] FileHash { get ; set; }
        public long CSID { get; set; }
        public string FileName { get ; set; }
        public string FileType { get ; set; }
        public long FileSize { get ; set; }
        public bool Deleted { get; set; }
    }
}
