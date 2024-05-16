using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AGOServer
{
    public class FolderInfo
    {
        private long nodeID;
        private long parentNodeID;
        private string folderName;
        private string iconUri;
        private string nodeType;
        private long childCount;

        public long NodeID { get => nodeID; set => nodeID = value; }
        public string Name { get => folderName; set => folderName = value; }
        public string IconUri { get => iconUri; set => iconUri = value; }
        public long ParentNodeID { get => parentNodeID; set => parentNodeID = value; }
        public string NodeType { get => nodeType; set => nodeType = value; }
        public long ChildCount { get => childCount; set => childCount = value; }
    }
}
