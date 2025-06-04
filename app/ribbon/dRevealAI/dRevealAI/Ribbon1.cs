using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

namespace dRevealAI
{
    public partial class Ribbon1 : Office.IRibbonExtensibility
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public string GetCustomUI(string RibbonID)
        {
            throw new NotImplementedException();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Hi");
        }
    }
}
