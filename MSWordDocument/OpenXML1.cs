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
    public partial class OpenXML1 : Form
    {
        private Microsoft.Office.Interop.Word.Application newWord = new Microsoft.Office.Interop.Word.Application();
        private Microsoft.Office.Interop.Word.Table currentTable;
        private string templatePath = @"C:\Temp\TryDocs\TryWordTable1.dotx";
        private string saveFileName = @"C:\Temp\TryDocs\TryWordTable1.docx";

        public OpenXML1()
        {
            InitializeComponent();

            newWord.Documents.Open(FileName: (templatePath)); //open as dotx document
            

            newWord.ActiveDocument.SaveAs(FileName: saveFileName, FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
            newWord.ActiveDocument.Close();
            newWord.Quit();
            //newWord.Documents.Add(path); //open as docx document
            //newWord.Documents.Open();
            //newWord.Visible = true;
        }


        private void Update_Click(object sender, EventArgs e)
        {
            SaveCurrentDocument();
        }

        private void SaveCurrentDocument()
        {
           

            using (var wpdocuments = WordprocessingDocument.Open(saveFileName, true))
            {
                Document doc = wpdocuments.MainDocumentPart.Document;
                Body body = wpdocuments.MainDocumentPart.Document.Body;

                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Insert text2222"));

                var tables = wpdocuments.MainDocumentPart.Document.Descendants<Table>().ToList();
                var tables2 = wpdocuments.MainDocumentPart.Document.Body.Elements<Table>().First();


                var tRowList = tables2.Descendants<TableRow>();

                foreach (TableRow tr in tRowList)
                {
                    var tCellList = tr.Descendants<TableCell>();

                    foreach (TableCell tc in tCellList)
                    {
                        string innerText = tc.InnerText;
                        Console.WriteLine(innerText);
                        //string innerXML = tc.InnerXml;
                        //Console.WriteLine(innerXML);
                    }
                }

                
                doc.Save();
            }
                        
        }

        private void OpenXML_FormClosed(object sender, FormClosedEventArgs e)
        {
            //newWord.ActiveDocument.Close(SaveChanges: Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            //newWord.Quit();
        }

        private void TableOne_Click(object sender, EventArgs e)
        {            

            using (var wpdocuments = WordprocessingDocument.Open(saveFileName, true))
            {                
                var tables = wpdocuments.MainDocumentPart.Document.Descendants<Table>().ToList();
                var tables2 = wpdocuments.MainDocumentPart.Document.Body.Elements<Table>().First();


                var tRowList = tables2.Descendants<TableRow>().ToList();
                TableRow tr = tRowList[1];

                for (int i = 0; i < 100; i++)
                {
                    var tCellList = tr.Descendants<TableCell>();
                    int rowIndex = i;
                    int colIndex = 0;

                    var newTr = new TableRow();
                    foreach (TableCell tc in tCellList)
                    {                         
                        string textValue = "Cell" + rowIndex.ToString() + "-" + colIndex.ToString();
                        string innerXML = tc.InnerXml;
                        innerXML = innerXML.Replace("{{Value}}", textValue);

                        TableCell newTc = new TableCell();
                        newTc.InnerXml = innerXML;

                        newTr.Append(newTc);
                        colIndex += 1;
                    }

                    tables2.Append(newTr);


                }

                tr.Remove();
                   
                

                Document doc = wpdocuments.MainDocumentPart.Document;
                doc.Save();
            }
        }

        private void ColumnMerge_Click(object sender, EventArgs e)
        {
            using (var wpdocuments = WordprocessingDocument.Open(saveFileName, true))
            {
                //TableCellProperties tableCellProperties = new TableCellProperties();
                //HorizontalMerge horizonMerge = new HorizontalMerge()
                //{
                //    Val = MergedCellValues.Restart
                //};
                //tableCellProperties.Append(horizonMerge);
                
                ////TableCellProperties tableCellProperties1 = new TableCellProperties();
                ////HorizontalMerge horizonMerge1 = new HorizontalMerge()
                ////{
                ////    Val = MergedCellValues.Continue
                ////};
                ////tableCellProperties1.Append(horizonMerge1);

                //TableCellProperties tableCellProperties3 = new TableCellProperties();
                //VerticalMerge verticalMerge3 = new VerticalMerge()
                //{
                //    Val = MergedCellValues.Restart
                //};
                //tableCellProperties3.Append(verticalMerge3);

                ////TableCellProperties tableCellProperties4 = new TableCellProperties();
                ////VerticalMerge verticalMerge4 = new VerticalMerge()
                ////{
                ////    Val = MergedCellValues.Continue
                ////};
                ////tableCellProperties4.Append(verticalMerge4);

                var tables = wpdocuments.MainDocumentPart.Document.Descendants<Table>().ToList();
                var tables2 = wpdocuments.MainDocumentPart.Document.Body.Elements<Table>().First();

                

                var tRowList = tables2.Descendants<TableRow>().ToList();
                TableRow tr = tRowList[1];

                int colCount = tr.Descendants<TableCell>().ToList().Count;

                for (int i = 0; i < 100; i++)
                {
                    var tCellList = tr.Descendants<TableCell>();
                    int rowIndex = i;
                    int colIndex = 0;

                    var newTr = new TableRow();
                    foreach (TableCell tc in tCellList)
                    {
                        string textValue = "Cell" + rowIndex.ToString() + "-" + colIndex.ToString();
                        string innerXML = tc.InnerXml;
                        innerXML = innerXML.Replace("{{Value}}", textValue);

                        TableCell newTc = new TableCell();
                        newTc.InnerXml = innerXML;

                        if (rowIndex == 5 && colIndex == 1)
                            newTc.Append(HorizontalMergeStartProperties());

                        if (rowIndex == 5 && colIndex == 2)
                            newTc.Append(HorizontalMergeContinueProperties());

                        if (rowIndex == 5 && colIndex == 3)
                            newTc.Append(HorizontalMergeContinueProperties());

                        if (rowIndex == 8 && colIndex == 1)
                            newTc.Append(VerticalMergeStartProperties());

                        if (rowIndex == 9 && colIndex == 1)
                            newTc.Append(VerticalMergeContinueProperties());

                        if (rowIndex == 10 && colIndex == 1)
                            newTc.Append(VerticalMergeContinueProperties());

                        if (rowIndex == 11 && colIndex == 1)
                            newTc.Append(VerticalMergeContinueProperties());

                        newTr.Append(newTc);
                        colIndex += 1;
                    }

                    tables2.Append(newTr);


                }

                tr.Remove();
                
                Document doc = wpdocuments.MainDocumentPart.Document;
                doc.Save();
            }
        }

        private TableCellProperties HorizontalMergeStartProperties()
        {
            TableCellProperties tableCellProperties = new TableCellProperties();
            HorizontalMerge horizonMerge = new HorizontalMerge()
            {
                Val = MergedCellValues.Restart
            };
            tableCellProperties.Append(horizonMerge);

            return tableCellProperties;
        }
        
        private TableCellProperties HorizontalMergeContinueProperties()
        {
            TableCellProperties tableCellProperties = new TableCellProperties();
            HorizontalMerge horizonMerge = new HorizontalMerge()
            {
                Val = MergedCellValues.Continue
            };
            tableCellProperties.Append(horizonMerge);

            return tableCellProperties;
        }

        private TableCellProperties VerticalMergeStartProperties()
        {
            TableCellProperties tableCellProperties = new TableCellProperties();
            VerticalMerge verticalMerge = new VerticalMerge()
            {
                Val = MergedCellValues.Restart
            };
            tableCellProperties.Append(verticalMerge);

            return tableCellProperties;
        }


        private TableCellProperties VerticalMergeContinueProperties()
        {
            TableCellProperties tableCellProperties = new TableCellProperties();
            VerticalMerge verticalMerge = new VerticalMerge()
            {
                Val = MergedCellValues.Continue
            };
            tableCellProperties.Append(verticalMerge);

            return tableCellProperties;
        }
    }
}
