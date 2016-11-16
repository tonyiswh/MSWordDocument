using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace MSWordDocument
{
    public partial class OpenXML : Form
    {
        private Microsoft.Office.Interop.Word.Application newWord = new Microsoft.Office.Interop.Word.Application();
        private Microsoft.Office.Interop.Word.Table currentTable;
        private string path = @"C:\Temp\TryDocs\TryWord.dotx";

        public OpenXML()
        {
            InitializeComponent();

            newWord.Documents.Open(FileName: (path)); //open as dotx document
            //newWord.Documents.Add(path); //open as docx document
            //newWord.Documents.Open();
            newWord.Visible = true;
        }


        private void Update_Click(object sender, EventArgs e)
        {
            SaveCurrentDocument();
        }

        private void SaveCurrentDocument()
        {
            //newWord.ActiveDocument.Save();
            string saveFileName = @"C:\Users\adm-tlin\Documents\TryWord22.docx";
            //string saveFileName = @"C:\Temp\TryDocs\TryWord22.docx";

            newWord.ActiveDocument.SaveAs(FileName: saveFileName, FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
            newWord.ActiveDocument.Close();
            newWord.Quit();

            WordprocessingDocument wpdocuments = WordprocessingDocument.Open(saveFileName, true);
            Body body = wpdocuments.MainDocumentPart.Document.Body;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Insert text"));

            var tables = wpdocuments.MainDocumentPart.Document.Descendants<TableProperties>().ToList();
            //var tables2 = wpdocuments.MainDocumentPart.Document.Body.Elements<Table>().First();
            var table = tables[0].Parent;
            if (table.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Table))
            {
               // Table aTable = (Table)table;

            }

            // Close the handle explicitly.
            wpdocuments.Close();

            

        }

        private void OpenXML_FormClosed(object sender, FormClosedEventArgs e)
        {
            //newWord.ActiveDocument.Close(SaveChanges: Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            //newWord.Quit();
        }
    }
}
