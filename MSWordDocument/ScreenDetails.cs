using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MSWordDocument
{
    public partial class ScreenDetails : Form
    {
        public ScreenDetails()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (var screen in System.Windows.Forms.Screen.AllScreens)
            {
                if (screen.Primary)
                {
                    var width = screen.WorkingArea.Width;
                    var height = screen.WorkingArea.Height;
                }
            }

           
        }

        private void ScreenDetails_Load(object sender, EventArgs e)
        {
           
            List<string> dbTypeSQLServerList = new List<string>()
            {
                "Binary",
                "Byte",
                "Boolean",
                "Currency",
                "Date",
                "DateTime",
                "Decimal",
                "Double",
                "Guid",
                "Int16",
                "Int32",
                "Int64",
                "Object",
                "SByte",
                "Single",
                "String",
                "Time",
                "UInt16",
                "UInt64",
            };

            DataTable dt = new DataTable();
            dt.Columns.Add("Column11");
            dt.Columns.Add("Column22");
            dt.Columns.Add("Column33");

            DataRow dr = dt.NewRow();
            dr["Column11"] = "Object";
            dr["Column22"] = "r12";
            dr["Column33"] = "r13";
            dt.Rows.Add(dr);

            DataRow dr2 = dt.NewRow();
            dr2["Column11"] = "Int64";
            dr2["Column22"] = "r22";
            dr2["Column33"] = "r23";
            dt.Rows.Add(dr2);

            
            dataGridView1.Rows.Clear();
            dataGridView1.AutoGenerateColumns = false;

            var col1 = (DataGridViewComboBoxColumn)dataGridView1.Columns["Column1"];
            col1.DataSource = dbTypeSQLServerList;
            //col1.DisplayMember = "Value";
            //col1.ValueMember = "Key";

            dataGridView1.Columns["Column1"].DataPropertyName = "Column11";
            dataGridView1.Columns["Column2"].DataPropertyName = "Column22";
            dataGridView1.Columns["Column3"].DataPropertyName = "Column33";
                       
            dataGridView1.DataSource = dt;
            dataGridView1.Refresh();

            SetTextColor();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ( e.ColumnIndex == 0 && e.RowIndex >= 0 )
            {
                var value = dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value;
                if (value.ToString().ToLower() ==  "object")
                {
                    dataGridView1.Rows[e.RowIndex].Cells["Column1"].Style.ForeColor = Color.Red;
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].Cells["Column1"].Style.ForeColor = Color.Black;
                }
            }
        }


        private void SetTextColor()
        {
            for(int r = 0 ; r < dataGridView1.Rows.Count - 1 ; r++)
            {
                var value = dataGridView1.Rows[r].Cells["Column1"].Value;
                if (value.ToString().ToLower() == "object")
                {
                    dataGridView1.Rows[r].Cells["Column1"].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[r].Cells["Column1"].Selected = false;
                }
            }

           

        }

        private void OpenPopup_Click(object sender, EventArgs e)
        {
            DataSetFields dsFields = new DataSetFields();
            dsFields.DatasetName = "First1";

            WPField wpField = new WPField();
            wpField.FieldName = "Field1";
            wpField.DataType = DbType.Int32;

            dsFields.Fields.Add(wpField);

            WPField wpField2 = new WPField();
            wpField2.FieldName = "Field2";
            wpField2.DataType = DbType.String;

            dsFields.Fields.Add(wpField2);

            WPField wpField3 = new WPField();
            wpField3.FieldName = "Field3";
            wpField3.DataType = DbType.String;

            dsFields.Fields.Add(wpField3);

            Popup popup = new Popup(dsFields);
            popup.ShowDialog();


        }
    }

    public class DataSetFields
    {
        private List<WPField> _fields = new List<WPField>();

        public List<WPField> Fields
        {
            get { return _fields; }
            set { _fields = value; }
        }

        public string DatasetName { get; set; }
    }

    public class WPField
    {
        public string FieldName { set; get; }

        public DbType DataType { get; set; }

        public bool Selected { get; set; }
    }

}
