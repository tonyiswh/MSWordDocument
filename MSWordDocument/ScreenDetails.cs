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
    }
}
