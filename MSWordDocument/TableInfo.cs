using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSWordDocument
{
    public class TableInfo
    {
        private List<RowInfo> _Rows = new List<RowInfo>();

        public string TableID { get; set; }

        public string TableType { get; set; }

        public string DataSource { get; set; }

        public List<RowInfo> Rows
        {
            get { return _Rows; }
            set { _Rows = value; }
        }



    }
}
