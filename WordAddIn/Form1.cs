using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Range range = Globals.ThisAddIn.Application.ActiveDocument.Content;
            object findText = "Test"; //string you want to find
            object oTrue = true;
            object oFalse = false;
            object oFindStop = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop;
            if (range.Find.Execute(ref findText, ref oTrue, ref oFalse, ref oTrue,
                 ref oFalse, ref oFalse, ref oTrue, ref oFindStop, ref oFalse))
            {
                range.Select();
                Globals.ThisAddIn.Application.Selection.Range.HighlightColorIndex = WdColorIndex.wdDarkYellow; //highlight the string which was found
            }

        }
    }
}
