using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSWordDocument
{
    public class RowInfo
    {
        private List<CellInfo> _Cells = new List<CellInfo>();

        public int RowIndex { get; set; }

        public string RowType { get; set; }

        public List<CellInfo> Cells
        {
            get { return _Cells; }
            set { _Cells = value; }
        }



    }
}
