using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class Rest_Node
    {
        public long ID { get; set; }
        public bool PermSeeContents { get; set; }
        public string Name { get; set; }
        public long ParentID { get; internal set; }
        public long Size { get; internal set; }
        public int Type { get; internal set; }
    }
}