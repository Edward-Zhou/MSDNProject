using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ExcelDoc
{
    partial class ActionsPaneControl1 : UserControl
    {
        public ActionsPaneControl1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("ok");
        }
    }
}
