using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Add_txt
{
    public partial class Add_text
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            var Form1 = new Form1();
            Form1.Show();
        }
    }
}
