using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSWordDocument
{
    public class TableBookmark
    {
        private List<string> _actionCodeList = new List<string>();
        public string TableID { get; set; }      

        public int RowIndex { get; set; }

        public int ColumnIndex { get; set; }

        public string BookmarkName { get; set; }

        public string BookmarkText { get; set; }

        public string DatasetName { get; set; }

        public string DataFieldName { get; set; }

        public List<string> ActionCodeList
        {
            get { return _actionCodeList; }
            set { _actionCodeList = value; }
        }

    }
}
