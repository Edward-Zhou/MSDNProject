using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Windows.Forms;

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
                Cell cell = GetCell(sheet, "C",1);
                var s = cell.LastChild;
                MessageBox.Show(s.InnerText);                
            }
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
