using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace AGOServer
{
    public class TableResponse
    {
        DataTable table;
        string[] columnHeaders;

        public DataTable Table { get => table; set => table = value; }
        public string[] ColumnHeaders { get => columnHeaders; set => columnHeaders = value; }
    }
}
