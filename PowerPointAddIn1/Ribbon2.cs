using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace PowerPointAddIn1
{
    public partial class Ribbon2
    {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Presentations presentations = ThisAddIn.instance.Application.Presentations;
            System.Diagnostics.Debug.WriteLine("!!! 1. " + presentations.Count.ToString());
            PowerPoint.Presentation presentation = presentations.Open(@"%file name of existing presentation here%");
            System.Diagnostics.Debug.WriteLine("!!! 2. " + presentations.Count.ToString());
            presentation.Close();
            System.Diagnostics.Debug.WriteLine("!!! 3. " + presentations.Count.ToString());
            Marshal.ReleaseComObject(presentation);
            System.Diagnostics.Debug.WriteLine("!!! 4. " + presentations.Count.ToString());
            Marshal.ReleaseComObject(presentations);
            button2_Click(null, null); 
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Presentations presentations = ThisAddIn.instance.Application.Presentations;
            System.Diagnostics.Debug.WriteLine("!!! button2_Click. " + presentations.Count.ToString());
            for (int i = 1; i <= presentations.Count; i++)
            {
                PowerPoint.Presentation presentation = presentations[presentations.Count];
                System.Diagnostics.Debug.WriteLine("!!! " + i.ToString() + ". " + presentation.FullName);
                Marshal.ReleaseComObject(presentation);
            }
            Marshal.ReleaseComObject(presentations);
            System.Diagnostics.Debug.WriteLine("!!! button2_Click. end");
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            app.Quit();
        }
    }
}
