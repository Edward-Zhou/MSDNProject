using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace OfficeAPI.Class.OpenXmlHelper
{
    public class GetCellValue
    {
        private SpreadsheetDocument document;

        public void getCell(string filePath)
        {
            using (document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                Worksheet sheet = wbPart.WorksheetParts.First().Worksheet;
                Cell cell = GetCell(sheet, "A",1);
                int columnIndex = GetColumnNumber(cell.CellReference.Value); //column name A
                Columns cs = sheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
                if (cs != null)
                {
                    IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Column> ic = cs.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(r => r.Min == columnIndex).Where(r => r.Max == columnIndex);
                    if (ic.Count() > 0)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Column c = ic.First();
                        double cc = c.Width;

                    }
                    else
                    {
                        
                        SheetFormatProperties p = sheet.Elements<SheetFormatProperties>().FirstOrDefault();
                        if (p == null)
                        {
                            double cc = 8.43;
                        }
                        else
                        {
                            if (p.DefaultColumnWidth == null)
                            {
                                double cc = 8.43;
                            }
                            else
                            {
                                double cc = p.DefaultColumnWidth;
                            }
                        }
                    }
                } 
                var s = cell.LastChild;
                MessageBox.Show(s.InnerText);                
            }
        }
        public static int GetColumnNumber(string name)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(name);
            name = match.Value;
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }
        private static Cell GetCell(Worksheet worksheet,
      string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();
        }


        // Given a worksheet and a row index, return the row.
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        } 
    }
}
