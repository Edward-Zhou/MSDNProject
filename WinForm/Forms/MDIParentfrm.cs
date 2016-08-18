using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinForm.Forms
{
    public partial class MDIParentfrm : Form
    {
        public MDIParentfrm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var form = new MDIChildfrm();
            form.ShowInTaskbar = false;
            form.MdiParent = this;
            form.Show();
        }
    }
}
