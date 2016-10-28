using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;

namespace MSWordDocument
{
    public partial class Table : Form
    {
        private Microsoft.Office.Interop.Word.Application newWord = new Microsoft.Office.Interop.Word.Application();
        private Microsoft.Office.Interop.Word.Table currentTable;
        private string path = @"C:\Temp\TryWord.docx";

        public Table()
        {
            InitializeComponent();

            newWord.Documents.Open(FileName: (path));
            newWord.Visible = true;
        }

        private void AddTable_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Range range = newWord.Selection.Range;
            Microsoft.Office.Interop.Word.Table table = newWord.ActiveDocument.Tables.Add(range, 10, 4);
            table.Range.Font.Size = 12;
            //table.set_Style("Light Shading - Accent 3");
            table.set_Style("Table Grid 8");
            //table.set_Style("Light List - Accent 5");

            currentTable = table;

            table.Rows.Add();


        }

        private void FillRows_Click(object sender, EventArgs e)
        {
            int row = 0, column = 0;
            //find selected cell indexes in a table 
            if (newWord.Selection.Information[Microsoft.Office.Interop.Word.WdInformation.wdWithInTable] == true )
            {
                row = newWord.Selection.Cells[1].RowIndex;
                column = newWord.Selection.Cells[1].ColumnIndex;
            }

            //Find a selected table
            foreach (Microsoft.Office.Interop.Word.Table tempTable in newWord.ActiveDocument.Tables)
            {
                if (newWord.Selection.Range.InRange(tempTable.Range))
                {
                    currentTable = tempTable;
                }
            }

            newWord.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1, Microsoft.Office.Interop.Word.WdMovementType.wdMove);
            //MoveDown not working
            //newWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdRow, 1, Microsoft.Office.Interop.Word.WdMovementType.wdMove);

            currentTable.Rows[row + 1].Cells[column + 1].Range.Select();
            newWord.Selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart);

            //move cursor out of table
            currentTable.Select();
            newWord.Selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            Microsoft.Office.Interop.Word.Bookmarks tempBookmarks = currentTable.Rows[row + 1].Cells[column + 1].Range.Bookmarks;

           


            foreach (Microsoft.Office.Interop.Word.Bookmark tempBookmark in tempBookmarks)
            {
                string name = tempBookmark.Name;
                string text = tempBookmark.Range.Text.ToString();
            }

            
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {            
            newWord.ActiveDocument.Close(SaveChanges: Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges );
            newWord.Quit();
        }

        private void TitleBookmarks_Click(object sender, EventArgs e)
        {

        }

        private void AddRow_Click(object sender, EventArgs e)
        {
            int rowCount = currentTable.Rows.Count;
            Microsoft.Office.Interop.Word.Range range = currentTable.Rows[rowCount].Cells[1].Range;

            currentTable.Rows.Add();
                      
            

            


        }
    }
}

////Get the Word range from the form's point location 
//Microsoft.Office.Interop.Word.Range range = (Microsoft.Office.Interop.Word.Range)Globals.ThisAddIn.Application.ActiveWindow.RangeFromPoint(x, y);
////Insert a dummy details table for the selected order
//Word.Table table = this.Application.ActiveDocument.Tables.Add(range, 4, 4);
//table.Range.Font.Size = 8;
//            table.set_Style("Table Grid 8");
//            table.Rows[1].Cells[1].Range.Text = "Order Details";
//            table.Rows[1].Cells[2].Range.Text = "Order Details";
//            table.Rows[1].Cells[3].Range.Text = "Order Details";
//            table.Rows[1].Cells[4].Range.Text = "Order Details";
//            for (int i = 2; i< 5; i++)
//			{
//                for (int j = 1; j< 5; j++)
//                {
//                    table.Rows[i].Cells[j].Range.Text = data.ToString();    
//                }

//			}

//set pTable = activedocument.Bookmarks("mybm").Range.Tables(1)
//activedocument.Bookmarks(oBookmark).Range.Text = strText
//newWord.ActiveDocument.Bookmarks[1].Column

//ActiveDocument.Bookmarks("MyBkMark").Select
//word_app.Selection.MoveRight(Word.WdUnits.wdCell, 1, Word.WdMovementType.wdMove)
//IEnumerator bookMarks = word_app.ActiveDocument.Bookmarks.GetEnumerator()
//While(bookMarks.MoveNext)
//Word.Bookmark book = bookMarks.Current
//string   name = book.Name.ToString();
//string text = book.Range.Text.ToString();