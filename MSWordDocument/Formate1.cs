using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MSWordDocument
{
    public partial class Formate1 : Form
    {
        public Formate1()
        {
            InitializeComponent();
        }

        private void btnNumber_Click(object sender, EventArgs e)
        {
            string cultureName = "fr-FR";
            CultureInfo culture = new CultureInfo(cultureName);
            String result = String.Format(culture, "Population {0:N}, Area {1:N3} sq. feet", 12212.22, 121288.5566);
                                    
        }
    }
}
