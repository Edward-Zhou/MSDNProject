using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane customTaskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public void AddAllTaskPanes()
        {
            PowerPoint.DocumentWindows windows = Globals.ThisAddIn.Application.Windows;
            if (windows.Count > 0)
            {
                customTaskPane = this.CustomTaskPanes.Where(c => c.Title == Globals.ThisAddIn.Application.ActiveWindow.Caption).FirstOrDefault();
                if (customTaskPane == null)
                {
                    customTaskPane = this.CustomTaskPanes.Add(new UserControl1(), Globals.ThisAddIn.Application.ActiveWindow.Caption, Globals.ThisAddIn.Application.ActiveWindow);
                }
                customTaskPane.Visible = true;
            }
        }
        public void AddAllTaskPanesWindows()
        {
            PowerPoint.DocumentWindows windows = Globals.ThisAddIn.Application.Windows;
            if (windows.Count > 0)
            {
                for (int i = 1; i <= windows.Count; i++)
                {
                    PowerPoint.DocumentWindow window = windows[i];
                    customTaskPane = this.CustomTaskPanes.Add(new UserControl1(), "My User Control", window);
                    customTaskPane.Visible = true;
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
