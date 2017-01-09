using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelDoc
{
    public partial class ThisWorkbook
    {
        private ActionsPaneControl1 ap = new ActionsPaneControl1();
        
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
           //this.ActionsPane.Controls.Add(ap);
           ////this.ActionsPane.StackOrder = Microsoft.Office.Tools.StackStyle.FromBottom;
           //this.Application.CommandBars["Task Pane"].Position = Microsoft.Office.Core.MsoBarPosition.msoBarBottom;
            Roller rooler=new Roller();
            ControlSite cs= Globals.Sheet1.Controls.AddControl(rooler, 0, 2, 1000, 30, "Roller1");
            Excel.Worksheet WS= this.Worksheets["sheet1"];
           
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
