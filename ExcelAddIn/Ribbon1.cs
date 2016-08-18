using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Threading;

namespace ExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            foreach (Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
            {
                RibbonDropDownItem rdd =this.Factory.CreateRibbonDropDownItem();
                rdd.Label = ws.Name;
                comboBox1.Items.Add(rdd);
            }
        }

        private void CreateShape_Click(object sender, RibbonControlEventArgs e)
        {

            Microsoft.Office.Interop.Excel.Shape btn2 = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddFormControl(Microsoft.Office.Interop.Excel.XlFormControl.xlButtonControl, 150, 5, 150, 22);
            btn2.ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset1;
            btn2.Name = "Update";
        }

        private void ExcelTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            string fileName = @"C:\Users\v-tazho\Desktop\Test.xlsx";
            string TemplateFileLocation = Path.GetFullPath(fileName);
            Worksheet wsadd = Globals.ThisAddIn.Application.ActiveSheet;
            if (File.Exists(fileName))
            {
                //Worksheet newWorkSheet = Globals.ThisAddIn.Application.Worksheets.Add(Missing.Value, Missing.Value, 1, fileName); //(WorkSheet)Globals.ThisAddin.Aplication.Worksheets.Add(Missing.Value, Missing.Value, 1, TemplateFileLocation);
                //Worksheet newWorkSheet = Globals.ThisAddIn.Application.Worksheets.Add(Type:fileName); 
                Workbook wb = Globals.ThisAddIn.Application.Workbooks.Open(@"C:\Users\v-tazho\AppData\Roaming\Microsoft\Templates\Test.xltx");
                Worksheet ws = wb.Worksheets[1];
                ws.Copy(After: wsadd);
            }
        }

        private void Classbtn_Click(object sender, RibbonControlEventArgs e)
        {
            LOCK_CELLS_PROTECT_CONTENTS(Globals.ThisAddIn.Application.ActiveSheet);
        }
        public void LOCK_CELLS_PROTECT_CONTENTS(Excel.Worksheet ws)
        {
            //  Locking Cells in ProtectContents
            MessageBox.Show("ok");
        }

        private void CopyPivotTable_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheetDestination = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1];//this.ActiveWorkbook.Worksheets[1];
            ////worksheetDestination.Range["A1"].Value = "Some Value"; // This is added so that when opening the other workbook this does not close

            //Workbook xlBook = (Workbook)worksheetDestination.Parent;
            Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application; //(Microsoft.Office.Interop.Excel.Application)xlBook.Parent;

            Workbook workbookOrg = xlApp.Workbooks.Open(@"C:\Users\v-tazho\Desktop\Test.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            System.Threading.Thread.Sleep(5000);
            Worksheet worksheetOrg = (Worksheet)workbookOrg.Worksheets["Sheet3"];

            worksheetOrg.Copy(Type.Missing, worksheetDestination);

            workbookOrg.Close();
        }

        private void CallMacro_Click(object sender, RibbonControlEventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application oExcel = Globals.ThisAddIn.Application;
            oExcel.Visible = true;
            Excel.Workbooks oBooks = oExcel.Workbooks;
            Excel._Workbook oBook = null;
            oBook = oBooks.Open(@"C:\Users\v-tazho\Desktop\Book4.xlsm", oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            // Run the macros.
            RunMacro(oExcel, new Object[] { "copyExcel" });
          
            // Quit Excel and clean up.
            //oBook.Close(false, oMissing, oMissing);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
            //oBook = null;
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
            //oBooks = null;
            //oExcel.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
            //oExcel = null;

        }
        private void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run",
                System.Reflection.BindingFlags.Default |
                System.Reflection.BindingFlags.InvokeMethod,
                null, oApp, oRunArgs);
        }

        private void qryTables_Click(object sender, RibbonControlEventArgs e)
        {
            //string Connection = @"D:\OfficeDev\Excel\T2.csv";
            Excel.QueryTable m_objQryTable = (Excel.QueryTable)Globals.ThisAddIn.Application.ActiveSheet.QueryTables.Add(@"TEXT;D:\OfficeDev\Excel\T1.csv", Globals.ThisAddIn.Application.Range["A1"], Missing.Value);
        }

        private void WorkBookSize_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application eApp = Globals.ThisAddIn.Application;
            eApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
            //get screen width and heigth
            double screenWidth = Screen.PrimaryScreen.Bounds.Width;
            double screenHeight = Screen.PrimaryScreen.Bounds.Height;
            //set excel applciation width and height
            eApp.ActiveWindow.Left = 178.75;
            eApp.ActiveWindow.Top = 251.5;
            eApp.ActiveWindow.Width = 481.5;
            eApp.ActiveWindow.Height = 266.25;
        }

        private void SelectUsedRange_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet ew = Globals.ThisAddIn.Application.ActiveSheet;
            ew.UsedRange.Select();
        }

        private void ExcelCopybtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application ea = Globals.ThisAddIn.Application;
            Excel.Worksheet ew= ea.ActiveWorkbook.Worksheets["Sheet1"];
            ew.Range["A1"].Copy();
            Excel.Range source = ew.Range["A1"];
            Excel.Worksheet ew2 = ea.ActiveWorkbook.Worksheets["Sheet2"];
            ew2.Range["A1"].PasteSpecial();
            Excel.Range target = ew2.Range["A1"];
            target.ColumnWidth = source.ColumnWidth;
            target.RowHeight = source.RowHeight;
        }

        private void ShowTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            OnBtnSourceAssociation(e.Control);
        }
        public void OnBtnSourceAssociation(IRibbonControl control)
        {
            Globals.ThisAddIn.OpenPanelSourceAssociation();
        }

        private void DeleteRange_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet sh = Globals.ThisAddIn.Application.Worksheets["Sheet1"];
            ((_Worksheet)sh).Activate();
            for (int i = sh.Protection.AllowEditRanges.Count; i >= 1; i--)
            {                
                sh.Unprotect("1234");
                sh.Protection.AllowEditRanges[i].Delete();
                sh.Protect("123");   
            }
        }

        private void AllowEditRangesbtn_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            
            sh.Protection.AllowEditRanges.Add("Protect",sh.get_Range("A1:A4"),"123");
            sh.Protect("123");
        }

        private void FilePre_Click(object sender, RibbonControlEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Typical Files(*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm|All files (*.*)|*.*";

            var res = openFileDialog.ShowDialog();

            if (res == DialogResult.OK)
            {
                MessageBox.Show(string.Format("{0} {1}", "File to open:", openFileDialog.FileName));
                //do somethink to open workbook
            }
            else
            {
                MessageBox.Show("Canceled..");
            }
        }

        private void ShapeSave_Click(object sender, RibbonControlEventArgs e)
        {
          Worksheet ws=  Globals.ThisAddIn.Application.ActiveSheet;
          ws.Shapes.Item(1).CopyPicture();
          ws.Range["A1"].Select();
          ws.Paste();
        }
        private MyUserControl myUserControl1;
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane1;
        //get the CustomTaskPane, and show or hide it
        private void CustomTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane myCustomTaskPaneShowHide = Globals.ThisAddIn.myCustomTaskPaneShowHide;
            if (myCustomTaskPaneShowHide.Visible)
            {
                myCustomTaskPaneShowHide.Visible = false;
            }
            else
            {
                myCustomTaskPaneShowHide.Visible = true;
            }
        }

        private void AddTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            myUserControl1 = new MyUserControl();
            //Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, "My Task Pane", Globals.ThisAddIn.Application.ActiveWindow);
            Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, Globals.ThisAddIn.Application.ActiveWorkbook.Name);


        }

        private void AddSecondTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            myUserControl1 = new MyUserControl();
            Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, "My Task Pane", Globals.ThisAddIn.Application.ActiveWindow);

        }

        private void TaskCount_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane ctp = Globals.ThisAddIn.CustomTaskPanes.Where(c => c.Title == Globals.ThisAddIn.Application.ActiveWorkbook.Name).FirstOrDefault();
            int i = Globals.ThisAddIn.CustomTaskPanes.Count();
            MessageBox.Show(ctp.Title);
        }

        private void ChangeDataLabel_Click(object sender, RibbonControlEventArgs e)
        {
            Series series = chart.FullSeriesCollection(1);            
            DataLabels db = series.DataLabels();
            db.Select();
            db.Format.AutoShapeType = MsoAutoShapeType.msoShapeOctagon;
        }
        Chart chart;
        private void AddDataLabel_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
            // Add chart.
            var charts = worksheet.ChartObjects() as
                Microsoft.Office.Interop.Excel.ChartObjects;
            var chartObject = charts.Add(60, 10, 300, 300) as
                Microsoft.Office.Interop.Excel.ChartObject;
            chart = chartObject.Chart;

            // Set chart range.
            var range = worksheet.get_Range("A1", "B3");
            chart.SetSourceData(range);

            // Set chart properties.
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterSmooth;
            chart.ChartWizard(Source: range,
                Title: "graphTitle",
                CategoryTitle: "xAxis",
                ValueTitle: "yAxis");
            chart.SetElement(MsoChartElementType.msoElementDataLabelTop);
        }

        private void DataLabelPosition_Click(object sender, RibbonControlEventArgs e)
        {
            Series series = chart.FullSeriesCollection(1);
            DataLabels db = series.DataLabels();
            //db.Select();
            //db.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionBelow;
            foreach (DataLabel dl in db)
            { 
                //MessageBox.Show(dl.Top.ToString());
                dl.Top=dl.Top+100;
            }
        }

        private void UnSaveBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveCell.Value = "Test";
            Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(@"D:\OfficeDev\Excel\201602\20160219.xlsx");
        }

        private void PivotFilter_Click(object sender, RibbonControlEventArgs e)
        {
              Globals.ThisAddIn.CreatePivotTable();
        }

        private void TopFilter_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.FilterPivotTable();
        }

        private void showForm_Click(object sender, RibbonControlEventArgs e)
        {
            var thread = new Thread(() =>
            {
                Form1 f = new Form1();
                f.ShowDialog();
            });
            thread.Start();
            //Form1 f = new Form1();
            //f.ShowDialog();
        }




		
    }
}
