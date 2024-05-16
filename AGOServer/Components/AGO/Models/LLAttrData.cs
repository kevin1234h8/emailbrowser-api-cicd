using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AGOServer
{
    public class LLAttrData
    {
        private long defID;
        private int attrID;
        private long valInt;
        private float? valReal;
        private DateTime? valDate;
        private string valStr;
        private string valLong;

        public string ValLong { get => valLong; set => valLong = value; }
        public DateTime? ValDate { get => valDate; set => valDate = value; }
        public float? ValReal { get => valReal; set => valReal = value; }
        public long ValInt { get => valInt; set => valInt = value; }
        public int AttrID { get => attrID; set => attrID = value; }
        public long DefID { get => defID; set => defID = value; }
        public string ValStr { get => valStr; set => valStr = value; }
    }
}
