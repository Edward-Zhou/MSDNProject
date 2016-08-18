using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace UnitTestProject
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            using (ExcelLifeTimeManager excelManager = new ExcelLifeTimeManager())
            {
                Worksheet activeWorkSheet = excelManager.Excel.ActiveSheet as Worksheet;
                string findWhat = "DataToFind";
                string replaceWith = "DataToReplace";
                if (activeWorkSheet != null)
                {
                    string[,] data = new string[3, 3];

                    for (int outerIndex = 0; outerIndex < data.GetUpperBound(0); outerIndex++)
                    {
                        for (int innerIndex = 0; innerIndex < data.GetUpperBound(1); innerIndex++)
                        {
                            data[outerIndex, innerIndex] = findWhat;
                        }
                    }

                    Range rangeToWriteData = activeWorkSheet.Range["A1", "C3"];
                    rangeToWriteData.Value2 = data;
                    Range activeCell = excelManager.Excel.ActiveCell;
                    Range r = activeWorkSheet.Range[activeCell.Address];
                    activeWorkSheet.Range["A5"].Value2 = activeCell.Address;
                    r.Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true);
                    activeWorkSheet.Range["A8"].Value2 = activeCell.Address;
                    // Make sure active cell is A1
                    //Assert.IsTrue((activeCell.Row == 1) && (activeCell.Column == 1) && 
                    //    (excelManager.Excel.Cells.Count == 1));
                    //Assert.IsTrue((activeCell.Row == 1) && (activeCell.Column == 1) );
                    //activeWorkSheet.Range["A5"].Value2 = activeCell;
                    //activeCell.Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true);
                    //activeWorkSheet.Range["A10"].Value2 = activeCell;
                    // We replaced only the active cell. We expect next occurence, so nextOccurence should not be null.
                    //Range nextOccurence = activeWorkSheet.UsedRange.Find(findWhat, activeCell, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true);
                    // Below assert is failing, since Range.Replace is replacing all the instances of search data in the worksheet.
                    //Assert.IsNotNull(nextOccurence);
                }
            }

        }
        /// <summary>
        /// Utility class to manage Excel instance.
        /// </summary>
        private class ExcelLifeTimeManager : IDisposable
        {
            internal Application Excel { get; private set; }

            /// <summary>
            /// Creates instance of <see cref="ExcelLifeTimeManager"/> class.
            /// </summary>
            public ExcelLifeTimeManager()
            {
                Excel = new Application();
                Excel.Visible = true;
                Excel.Workbooks.Add();
            }

            /// <summary>
            /// Clean up the resources.
            /// </summary>
            public void Dispose()
            {
                //foreach (Workbook workbook in Excel.Workbooks)
                //{
                //    workbook.Close(SaveChanges: false);
                //}
                //Excel.Quit();
                //Excel = null;
            }
        }
    }
}
