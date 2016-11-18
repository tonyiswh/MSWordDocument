using NReco.PivotData;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MSWordDocument
{
    public partial class TryPivot : Form
    {
        public TryPivot()
        {
            InitializeComponent();
        }

        private void Pivot_Click(object sender, EventArgs e)
        {
            //NReco.PivotData
            //https://www.nrecosite.com/pivot_data_library_net.aspx
            //http://www.codeproject.com/Articles/22008/C-Pivot-Table


            string gSqlConnStr = "Data Source = DWS02;Initial Catalog = GUROOther; Integrated Security = True;";
            string strSQL = "select * from tblTransactions";
            var objConn = new SqlConnection(gSqlConnStr);
            objConn.Open();

            SqlCommand objcmd = new SqlCommand(strSQL, objConn);
            objcmd.CommandType = CommandType.Text;

            DataTable dt = new DataTable();
            using (SqlDataAdapter da = new SqlDataAdapter(objcmd))
            {
                da.Fill(dt);
            }

            objConn.Close();

            //dataGridView1.DataSource = dt;
            //dataGridView1.Show();

            var pivotData = new PivotData(new string[] { "chardate", "country"}, new SumAggregatorFactory("totalamount"), new DataTableReader(dt));


            var grandTotal = pivotData[Key.Empty, Key.Empty].Value;
            var subTotalFor29 = pivotData[Key.Empty, 29].Value;
            var allDimensionKeys = pivotData.GetDimensionKeys();

            List<object> rowDimensionKeys = allDimensionKeys[0].ToList();

            var pivotTable = new PivotTable(new[] { "chardate" }, new[] { "country"}, pivotData);

           
       

            int rowCount = pivotTable.RowKeys.Length;
            int colCount = pivotTable.ColumnKeys.Length;
            

            DataTable tableNew = new DataTable();
            tableNew.Columns.Add("chardate", typeof(object));

            for (int i = 0; i < colCount; i++)
            {
                string columnKey = pivotTable.ColumnKeys[i].ToString().Replace(" ", "").Replace("[", "").Replace("]","");
                tableNew.Columns.Add(columnKey, typeof(object));
            }

            for (int i = 0; i< rowCount; i++)
            {
                string rowKey = pivotTable.RowKeys[i].ToString().Replace(" ", "").Replace("[", "").Replace("]", "");
                List<object> rowValues = new List<object>();
                rowValues.Add(rowKey);
                for(int j= 0; j < colCount; j++)
                {
                    rowValues.Add(pivotTable[i, j].Value);
                }

                tableNew.Rows.Add(rowValues.ToArray());
            }


            dataGridView1.DataSource = tableNew;
            dataGridView1.Show();

            //object value = pivotTable[1, 2].Value;
        }
    }
}
