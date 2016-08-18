using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        private MyUserControl myUserControl1;
        private UserControl1 WPFUserControl;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        //define CustomTaskPane
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPaneShowHide;
        private Microsoft.Office.Tools.ActionsPane myCustomActionPane;
        private string sheetName;
        private Microsoft.Office.Interop.Excel.Worksheet ws;
        internal Helper helper;


        private CExcelCtrl Ctrl;
        private Excel.Application ExcelApp { get { return this.Application; } }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            helper = new Helper();
            myUserControl1 = new MyUserControl();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            myCustomTaskPane.Visible = true;
            //Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            //Ctrl = new CExcelCtrl(ExcelApp);

            //Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet();
            //ws.BeforeDoubleClick
           // Application.WorkbookOpen += Application_WorkbookOpen;

            //Application.SheetActivate += Application_SheetActivate;

            //Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
            //Application.WindowActivate += new Excel.AppEvents_WindowActivateEventHandler(Application_WindowActivate);
            //Application.SheetActivate+=Application_SheetActivate;
            System.Windows.Forms.MessageBox.Show("要显示的内容");
        }



        void Application_SheetChange(object Sh, Range Target)
        {

            for (int i = Globals.Ribbons.Ribbon1.comboBox1.Items.Count-1; i >= 0; i--)
            {
                Globals.Ribbons.Ribbon1.comboBox1.Items.RemoveAt(i);
            }
                foreach (Microsoft.Office.Interop.Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
                {
                    RibbonDropDownItem rdd = Globals.Ribbons.Ribbon1.Factory.CreateRibbonDropDownItem();
                    rdd.Label = ws.Name;
                    Globals.Ribbons.Ribbon1.comboBox1.Items.Add(rdd);
                }
            //iRibbonUI.ribbon.Invalidate();
            //Globals.Ribbons.Ribbon1.RibbonUI.Invalidate();
            //Globals.Ribbons.Ribbon1.RibbonUI.InvalidateControl("comboBox1");
        }

        private void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            Ribbon2 r = new Ribbon2();
            Microsoft.Office.Tools.CustomTaskPane ctp = Globals.ThisAddIn.CustomTaskPanes.Where(c => c.Title == Globals.ThisAddIn.Application.ActiveWorkbook.Name).FirstOrDefault();
            if (ctp == null)
            {
                myUserControl1 = new MyUserControl();
                myCustomTaskPaneShowHide = Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, Globals.ThisAddIn.Application.ActiveWorkbook.Name);

                //myCustomTaskPaneShowHide = Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, Globals.ThisAddIn.Application.ActiveWorkbook.Name, Globals.ThisAddIn.Application.ActiveWindow);
            }
            else
            {
                myCustomTaskPaneShowHide = ctp;
            }
            //myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes[1];
        }
        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            MessageBox.Show("Did you want to close");

        }
        private void Application_SheetActivate(object Sh)
        {
           
            iRibbonUI.ribbon.Invalidate();
            //for (int i = Globals.Ribbons.Ribbon1.comboBox1.Items.Count - 1; i >= 0; i--)
            //{
            //    Globals.Ribbons.Ribbon1.comboBox1.Items.RemoveAt(i);
            //}
            //foreach (Microsoft.Office.Interop.Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
            //{
            //    RibbonDropDownItem rdd = Globals.Ribbons.Ribbon1.Factory.CreateRibbonDropDownItem();
            //    rdd.Label = ws.Name;
            //    Globals.Ribbons.Ribbon1.comboBox1.Items.Add(rdd);
            //}
            //if (sheetName!=ws.Name)
            //{
            //    ws = Globals.ThisAddIn.Application.ActiveSheet;
            //    sheetName = ws.Name;
            //    MessageBox.Show(sheetName);

            //}
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            ws =Globals.ThisAddIn.Application.ActiveSheet;
            sheetName = ws.Name;
            MessageBox.Show(sheetName);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }
        public void OpenPanelSourceAssociation()
        {
            myCustomTaskPane.Width = 275;
            OpenPanel(myCustomTaskPane);
        }
        private void OpenPanel(Microsoft.Office.Tools.CustomTaskPane ActpPanel)
        {
            ActpPanel.Visible = true;
        }
        private void Application_WorkbookBeforeSave(Excel.Workbook workbook, bool saveAs, ref bool cancel)
        {
            Console.Write(saveAs);
            //Trace.WriteLine("Application_WorkbookBeforeSave saveAs = " + saveAs);
        }
        Excel.PivotTable pivotTable;
        internal void CreatePivotTable()
        {

            Excel.Worksheet ws = Application.ActiveSheet;
            long lastColumn = ws.Cells[1, ws.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            long lastRow = ws.Cells[ws.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;
            Excel.Range dataRange = ws.Range[ws.Cells[1.1], ws.Cells[lastRow, lastColumn]];
            string address = dataRange.Address;
            //Excel.Range r = ws.UsedRange;
            //r.Select();
            Excel.Worksheet newSheet = Application.ActiveWorkbook.Worksheets.Add();
            Excel.PivotCache pivotCache = Application.ActiveWorkbook.PivotCaches().Create(SourceType: Excel.XlPivotTableSourceType.xlDatabase, SourceData: dataRange);//R1C1:R394C25
            pivotTable = pivotCache.CreatePivotTable(TableDestination: newSheet.Cells[1, 1], TableName: "TFSARReport5");
        }

        internal void FilterPivotTable()
        {
            pivotTable.PivotFields("Y").PivotFilters.Add(XlPivotFilterType.xlTopCount, pivotTable.PivotFields("Sum of Y"), 5);

        }

        //Ribbon1 r1 = new Ribbon1();

        //protected override Microsoft.Office.Tools.Ribbon.IRibbonExtension[] CreateRibbonObjects()
        //{
        //    r1 = new Ribbon1();
        //    return new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { r1 };
        //}  
        private Ribbon3 iRibbonUI;
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            iRibbonUI = new Ribbon3();
            return iRibbonUI;
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

    class CExcelCtrl
    {
        private Excel.Application ExcelApp { get; set; }
        public Excel.Workbook ExcelWbk { get; set; }
        private System.Windows.Forms.Timer STimer;
        private string wbName;
        private int count;

        public CExcelCtrl(Excel.Application app)
        {
            ExcelApp = app;
            ExcelWbk = ExcelApp.ActiveWorkbook;
            wbName = ExcelWbk.Name;
            count = 0;
            STimer = new System.Windows.Forms.Timer();

            STimer.Interval = 10000;
            STimer.Enabled = true;
            STimer.Tick += new EventHandler(STimer_EventProcessor);
        }
        public void STimer_EventProcessor(object sender, EventArgs e)
        {
            try
            {
                STimer.Stop();
                if (wbName == ExcelWbk.Name)//exception with message:Exception from HRESULT: 0x800401A8
                    count++;

            }
            catch (Exception exc)
            {
                //it is strange, it is work for a few times and failure after some times when I debug this add-in
            }

            STimer.Start();

        }
    }
}
