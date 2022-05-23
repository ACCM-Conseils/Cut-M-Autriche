using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CUT_M
{
    public partial class Warning : Form
    {
        public Warning()
        {
            InitializeComponent();

            webBrowser1.Navigate(ut_xml.ValueXML(@".\CUT-M.xml", "Avertissement"));
        }
    }
}
