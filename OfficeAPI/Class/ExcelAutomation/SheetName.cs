using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAPI.Class.ExcelAutomation
{
    class SheetName
    {
        public static void TriggerExcel2010Bug()
        {
            //string fileLocation = AppDomain.CurrentDomain.BaseDirectory + @"Sample Pivot Chart.xlsx";
            string fileLocation = @"C:\Users\v-tazho\Desktop\test.xlsx";
            // Open the interop
            Application xlApplication = new Application();

            // Open our example workbook, and select our example worksheet
            Workbook workbook = xlApplication.Workbooks.Open(fileLocation);
            Worksheet worksheetWithCharts = workbook.Worksheets["Sheet1"];
            xlApplication.Visible = true;

            // Rename the sheet to a new name
            worksheetWithCharts.Name = "a new name";

            // Copy. Code will fail here in Excel 2010.
            worksheetWithCharts.Copy(workbook.Worksheets[1], Type.Missing);

            // Clean up
            workbook.Close(false, Type.Missing, Type.Missing);
            xlApplication.Quit();
        }
        public static void AddConnectionString()
        {
            //string fileLocation = AppDomain.CurrentDomain.BaseDirectory + @"Sample Pivot Chart.xlsx";
            string fileLocation = @"D:\OfficeDev\Excel\Excel.xlsx";
            // Open the interop
            Application xlApplication = new Application();

            // Open our example workbook, and select our example worksheet
            Workbook workbook = xlApplication.Workbooks.Open(fileLocation);
            Worksheet worksheetWithCharts = workbook.Worksheets["Sheet1"];
            xlApplication.Visible = true;
            Microsoft.Office.Interop.Excel.QueryTable m_objQryTable = (Microsoft.Office.Interop.Excel.QueryTable)worksheetWithCharts.QueryTables.Add(@"", worksheetWithCharts.get_Range("A1", Missing.Value), Missing.Value);    
        }
    }
}
