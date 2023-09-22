using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Resources;
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

            ResourceManager res_man = new ResourceManager("CUT_M.Lang_" + System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName, Assembly.GetExecutingAssembly());

            button1.Text = res_man.GetString("Consigne");

            webBrowser1.Navigate(ut_xml.ValueXML(@".\CUT-M.xml", "Avertissement"));
        }
    }
}
