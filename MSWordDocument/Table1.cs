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
using System.Dynamic;
using Microsoft.Office.Interop.Word;

namespace MSWordDocument
{
    public partial class Table1 : Form
    {
        private Microsoft.Office.Interop.Word.Application newWord = new Microsoft.Office.Interop.Word.Application();
        private Microsoft.Office.Interop.Word.Table currentTable;
        private string path = @"C:\Temp\TryDocs\TryWord.dotx";

        public Table1()
        {
            InitializeComponent();

            newWord.Documents.Open(FileName: (path));
            //newWord.Documents.Add(path);
            //newWord.Documents.Open();
            newWord.Visible = true;
        }

        private void AddTable_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Range range = newWord.Selection.Range;
            Microsoft.Office.Interop.Word.Table table = newWord.ActiveDocument.Tables.Add(range, 10, 4, Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord8TableBehavior, Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow);
            table.Range.Font.Size = 12;
           
            //table.set_Style("Light Shading - Accent 3");
            table.set_Style("Table Grid 8");
            //table.set_Style("Light List - Accent 5");            
            table.ID = "My Table ID";
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

            //Fin a table by ID
            foreach (Microsoft.Office.Interop.Word.Table tempTable in newWord.ActiveDocument.Tables)
            {
                if (tempTable.ID.ToLower() == "my table id")
                {
                    tempTable.Select();
                    break;
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

            //Find bookmark in a table
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
            GetCurrentTable();

            int rowCount = currentTable.Rows.Count;
            Microsoft.Office.Interop.Word.Range range = currentTable.Rows[rowCount].Cells[1].Range;

            
            currentTable.Rows.Add(); //add a last row


            dynamic ed = new ExpandoObject();

            ed.Name = "asdfs";
            ed.Address = "asdfsdf";
            
            ExpandoObject edo = new ExpandoObject();
            ((IDictionary<string, object>)edo).Add("", "");
        }

        private void CopyTable_Click(object sender, EventArgs e)
        {
            foreach (Microsoft.Office.Interop.Word.Table tempTable in newWord.ActiveDocument.Tables)
            {
                if (newWord.Selection.Range.InRange(tempTable.Range))
                {
                    currentTable = tempTable;
                }
            }

            //currentTable.Range.Select();
            //currentTable.Range.Copy();
            //newWord.Selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
            ////newWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 2, Microsoft.Office.Interop.Word.WdMovementType.wdMove);
            //newWord.Selection.TypeParagraph();
            //newWord.Selection.TypeParagraph();
            //newWord.Selection.Range.Paste();

            currentTable.Range.Select();
            dynamic styleDyn = currentTable.get_Style();
            string styleName = styleDyn.NameLocal;
            string id = currentTable.ID;

            Microsoft.Office.Interop.Word.Range rangeText = currentTable.Range.FormattedText;
            string text = rangeText.XML;

            text = text.Replace("#aaa#", "New AAA");
            text = text.Replace("#bbb#", "New BBB");
            newWord.Selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
            newWord.Selection.TypeParagraph();
            newWord.Selection.TypeParagraph();
            //newWord.Selection.Range.FormattedText = rangeText;
           
            newWord.Selection.Range.InsertXML(text);            
            currentTable = newWord.Selection.Tables[1];
            currentTable.set_Style(styleName);
            currentTable.ID = id + "2";

            //string xml = rangeText.XML;
            //Console.Write(xml);
        }

        private void AddColumn_Click(object sender, EventArgs e)
        {
            float height = newWord.ActiveDocument.PageSetup.PageHeight - 50;
            float width = newWord.ActiveDocument.PageSetup.PageWidth - 100;

            GetCurrentTable();
            currentTable.Columns.Add(); //add a last column
            //currentTable.Columns.DistributeWidth();

            //Not working
            //width = width / currentTable.Columns.Count;
            //currentTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints;
            //currentTable.Columns.PreferredWidth = width;

            //currentTable.Columns[currentTable.Columns.Count].SetWidth(40, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustProportional);
            currentTable.Columns[currentTable.Columns.Count].SetWidth(40, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustSameWidth);
        }

        private void GetCurrentTable()
        {
            foreach (Microsoft.Office.Interop.Word.Table tempTable in newWord.ActiveDocument.Tables)
            {
                if (newWord.Selection.Range.InRange(tempTable.Range))
                {
                    currentTable = tempTable;
                }
            }
        }

        private void AddMoreRows_Click(object sender, EventArgs e)
        {
            GetCurrentTable();
            currentTable.AllowAutoFit = false;
           

            //Microsoft.Office.Interop.Word.WdViewType viewtype = newWord.ActiveWindow.View.Type;
            //bool pagination = newWord.Options.Pagination;
            //bool screenUpdating = newWord.ScreenUpdating;


            for (int i = 0; i < 100; i++)
            {
                int rowIndex = currentTable.Rows.Count - 2;

                currentTable.Rows.Add(currentTable.Rows[rowIndex]);
                currentTable.Rows[rowIndex].Cells[1].Range.Text = "Cell1 " + i.ToString();
                currentTable.Rows[rowIndex].Cells[2].Range.Text = "Cell2 " + i.ToString();
                currentTable.Rows[rowIndex].Cells[3].Range.Text = "Cell3 " + i.ToString();
                currentTable.Cell(rowIndex, 4).Range.Text = "Cell4 " + i.ToString();
            }

            //newWord.ActiveWindow.View.Type = viewtype;
            //newWord.Options.Pagination = pagination;
            //newWord.ScreenUpdating = screenUpdating;

            currentTable.Rows[currentTable.Rows.Count - 1].Cells[1].Range.Text = "Done";
        }

        private void SelectRow_Click(object sender, EventArgs e)
        {
            GetCurrentTable();

            currentTable.Rows[3].Select();
            string xml = newWord.Selection.Range.XML;
            Console.Write(xml);
        }

        private void ChackBox_Click(object sender, EventArgs e)
        {
            GetCurrentTable();

            string idtitle = currentTable.Title;            

            newWord.ActiveDocument.FormFields["Check2"].CheckBox.Value = true;

            /////////not working in select a bookmark than checkbox field 
            //Microsoft.Office.Interop.Word.Bookmark bookmark = newWord.ActiveDocument.Bookmarks["Check2"];            
            //bookmark.Range.Select();          
            //var conControl = bookmark.Range.Fields[1];
            //if (conControl.Type == Microsoft.Office.Interop.Word.WdFieldType.wdFieldFormCheckBox)
            //{
                
                
            //}

            newWord.ActiveDocument.FormFields["Check2"].CheckBox.Value = true;

            if (newWord.ActiveDocument.FormFields[1].Type == Microsoft.Office.Interop.Word.WdFieldType.wdFieldFormCheckBox)
            {
                newWord.ActiveDocument.FormFields[1].CheckBox.Value = true;
            }

            newWord.Selection.Move(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1);



            //Add a new checkbox and check it
            var newControl = newWord.ActiveDocument.FormFields.Add(newWord.Selection.Range, Microsoft.Office.Interop.Word.WdFieldType.wdFieldFormCheckBox);
            newControl.CheckBox.Value = true;

            
        }

        private void Dialog_Click(object sender, EventArgs e)
        {


            Microsoft.Office.Interop.Word.Dialog dialogBox = newWord.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogEditReplace];

            dialogBox.Show();
        }

        private void CurrentSelection_Click(object sender, EventArgs e)
        {
            string codeText = null;

            codeText = GetFieldName();

            Console.WriteLine(codeText);
            newWord.Application.StatusBar = "Code text: " + codeText;
        }

        public string GetFieldName()
        {
            string codeText = null;

            //get current cursor code text
            string fullText = newWord.Selection.Sentences[1].Text;

            string currentText = newWord.Selection.Range.Words[1].Text.Trim();

            //var col = newWord.Selection.get_Information(WdInformation.wdFirstCharacterColumnNumber);
            //var row = newWord.Selection.get_Information(WdInformation.wdEndOfRangeRowNumber);
            //var pos = newWord.Selection.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);

            if (currentText == "{{" || currentText == "{{}}")
            {
                return null;
            }

            if (currentText == "}}")
            {
                newWord.Selection.MoveLeft(WdUnits.wdCharacter, 1);
                currentText = newWord.Selection.Range.Words[1].Text.Trim();
                if (currentText == "}}" || currentText == "{{")
                {
                    return null;
                }
            }

            int index = fullText.IndexOf(currentText) + currentText.Length;
            string firstPart = fullText.Substring(0, index);
            int firstIndex = firstPart.LastIndexOf("{{", firstPart.Length);

            if (firstIndex < 0)
            {
                codeText = null;
            }
            else
            {
                firstIndex += 2;
                int secondIndex = fullText.IndexOf("}}", index);

                if (firstIndex > 0 && secondIndex > 0 && firstIndex < secondIndex)
                {
                    codeText = fullText.Substring(firstIndex, secondIndex - firstIndex).Trim();

                    if (codeText.Contains("{{") || codeText.Contains("}}"))
                    {
                        codeText = null;
                    }
                }
            }

            return codeText;
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


//Popup text box
//Ctrl + F9
//AutoTextList "word seen" \s NoStyle \t "text in the box"
//Alt + F9

//Range.ExportFragment
//Range.ImportFragment