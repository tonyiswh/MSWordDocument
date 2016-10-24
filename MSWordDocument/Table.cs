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
        private string path = @"C:\Temp\TryWord.docx";

        public Table()
        {
            InitializeComponent();

            newWord.Documents.Open(FileName: (path));
            newWord.Visible = true;
        }

        private void AddTable_Click(object sender, EventArgs e)
        {

        }

        private void FillRows_Click(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            newWord.ActiveDocument.Close();
            newWord.Quit();
        }

        private void TitleBookmarks_Click(object sender, EventArgs e)
        {

        }
    }
}
