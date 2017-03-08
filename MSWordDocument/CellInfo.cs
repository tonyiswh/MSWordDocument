using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSWordDocument
{
    public class CellInfo
    {
        private List<string> _DataFields = new List<string>();

        public int RowIndex { get; set; }

        public int ColumnIndex { get; set; }

        public bool IdGroupField { get; set; }

        public bool MergeRow { get; set; }

        public bool MergeColumn { get; set; }

        public List<string> DataFields
        {
            get { return _DataFields; }
            set { _DataFields = value; }
        }


    }
}
