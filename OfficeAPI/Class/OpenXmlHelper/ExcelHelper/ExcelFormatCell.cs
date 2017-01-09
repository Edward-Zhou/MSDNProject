using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Excel= OfficeAPI.Class.OpenXmlHelper.ExcelHelper;
//{
    //public class ExcelFormatCell
    //{
    //    public void FormatCell(string filePath)
    //    {
    //        //var ExcelObj = new Excel.Application();
    //        ////Excel.Workbook Book0 = ExcelObj.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
    //        //ExcelObj.ScreenUpdating = false;

    //        //Excel.Worksheet Worksheet0;

    //        using (FileStream fs = new FileStream(openFileDialog1.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    //        {
    //            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
    //            {
    //                WorkbookPart workbookPart = doc.WorkbookPart;

    //                Sheet theSheet = workbookPart.Workbook.Descendants<Sheet>().
    //                    Where(s => s.Name == "Sheet Overall").FirstOrDefault();

    //                if (theSheet == null)
    //                {
    //                    throw new ArgumentException("sheetName");
    //                }

    //                int sheetIndex = 0;
    //                foreach (WorksheetPart wsPart in workbookPart.WorksheetParts)
    //                {
    //                    WorkbookStylesPart wstylePart = workbookPart.WorkbookStylesPart;
    //                    Stylesheet ss = wstylePart.Stylesheet;

    //                    string sheetName = workbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex).Name;
    //                    if (sheetName != "Sheet Overall")
    //                    {
    //                        if (sheetName == null)
    //                        {
    //                            throw new ArgumentException("sheetName");
    //                        }

    //                        using (StreamWriter sw = new StreamWriter(@File1, true, Encoding.Unicode))
    //                        {
    //                            sw.WriteLine("Sheet name: {0}", sheetName);
    //                        }

    //                        var cells = wsPart.Worksheet.Descendants<Cell>();

    //                        foreach (Cell cell in cells)
    //                        {
    //                            int i;
    //                            //if (string.IsNullOrEmpty(cell.StyleIndex.Value.ToString()))
    //                            if (cell.StyleIndex != null)
    //                            {
    //                                i = Convert.ToInt32(cell.StyleIndex.Value);

    //                                //Column cellColumn = wsPart.Worksheet.Descendants<Column>().ElementAt(GetColumnIndex(cell.CellReference.Value));

    //                                DocumentFormat.OpenXml.Spreadsheet.CellFormat cellFormat = ss.Descendants<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt(i);

    //                                int j = Convert.ToInt32(cellFormat.FontId.Value);
    //                                int k = Convert.ToInt32(GetColumnIndex(cell.CellReference.Value));
    //                                DocumentFormat.OpenXml.Spreadsheet.Font font = ss.Descendants<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt(j);
    //                                FontSize fontsize = font.FontSize;

    //                                //using (StreamWriter sw = new StreamWriter("D:\\cellValue.txt", true, Encoding.Unicode))
    //                                foreach (Column cellColumn in wsPart.Worksheet.Descendants<Column>())
    //                                {
    //                                    if (cellColumn.Min.Value == j)
    //                                    {
    //                                        var cellWidth = cellColumn.Width.Value;
    //                                        using (StreamWriter sw = new StreamWriter(@File1, true, Encoding.Unicode))
    //                                        {
    //                                            sw.WriteLine("Cell contents: {0}", GetCellValue(cell, workbookPart));
    //                                            sw.WriteLine("Cell width: {0}", cellWidth.ToString());
    //                                            sw.WriteLine("Cell font: {0}", fontsize.Val.ToString());
    //                                        }
    //                                    }

    //                                }
    //                            }
    //                        }

    //                    }
    //                    sheetIndex++;
    //                }
    //            }
    //        }
    //    }
    //}
//}
