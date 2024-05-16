using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AGOServer
{
    public class FolderSearchInfo
    {
        private long nodeID;
        private long parentNodeID;
        private string folderName;
        private long childCount;
        private string fullPath;

        public long NodeID { get => nodeID; set => nodeID = value; }
        public string Name { get => folderName; set => folderName = value; }
        public long ParentNodeID { get => parentNodeID; set => parentNodeID = value; }
        public long ChildCount { get => childCount; set => childCount = value; }
        public string FullPath { get => fullPath; set => fullPath = value; }
    }
}
