using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Windows.Forms;
using System.IO.Packaging;
using System.IO;
using System.Xml;

namespace OfficeAPI.Class.OpenXmlHelper
{
    public class setCellValue
    {
        private SpreadsheetDocument document;

        public void setCell(string filePath)
        {
            using (document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                
                Worksheet sheet = wbPart.WorksheetParts.First().Worksheet;

                //ChangeWorksheetPart(sheet.WorksheetPart);
                ChangeWorksheetPart1(sheet.WorksheetPart);
                //SheetData sheetData1 = sheet.GetFirstChild<SheetData>();
                ////Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, DyDescent = 0.15D };
                ////Cell cell = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U };
                ////CellValue cellValue = new CellValue();
                //////cellValue.Text = DateTime.Now.ToOADate().ToString();
                ////cellValue.Text = "04.01.2013";
                ////cell.Append(cellValue);
                ////row1.Append(cell);
                ////sheetData1.Append(row1);
                ////number
                //Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, DyDescent = 0.25D };

                //Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U };
                //CellValue cellValue1 = new CellValue();
                //cellValue1.Text = "1000.00";

                //cell1.Append(cellValue1);

                //row1.Append(cell1);
                //sheetData1.Append(row1);

            }
        }
        public void ChangeWorksheetPart(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet1.MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" };

            SheetViews sheetViews1 = worksheet1.GetFirstChild<SheetViews>();
            SheetFormatProperties sheetFormatProperties1 = worksheet1.GetFirstChild<SheetFormatProperties>();
            SheetData sheetData1 = worksheet1.GetFirstChild<SheetData>();

            SheetView sheetView1 = sheetViews1.GetFirstChild<SheetView>();

            Selection selection1 = new Selection() { ActiveCell = "C3", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C3" } };
            sheetView1.Append(selection1);
            sheetFormatProperties1.DyDescent = 0.25D;

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, DyDescent = 0.25D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U };
            CellValue cellValue1 = new CellValue();
            cell1.DataType = new EnumValue<CellValues>(CellValues.Number);
            cellValue1.Text = "1000.00";        
                              

            cell1.Append(cellValue1);

            row1.Append(cell1);
            sheetData1.Append(row1);
        }
        public void ChangeWorksheetPart1(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;
            worksheet1.RemoveNamespaceDeclaration("x");

            SheetViews sheetViews1 = worksheet1.GetFirstChild<SheetViews>();
            SheetData sheetData1 = worksheet1.GetFirstChild<SheetData>();

            SheetView sheetView1 = sheetViews1.GetFirstChild<SheetView>();

            Selection selection1 = sheetView1.GetFirstChild<Selection>();
            selection1.ActiveCell = "G8";
            selection1.SequenceOfReferences = new ListValue<StringValue>() { InnerText = "G8" };

            Row row1 = sheetData1.GetFirstChild<Row>();

            Cell cell1 = row1.GetFirstChild<Cell>();
            cell1.DataType = null;

            CellValue cellValue1 = cell1.GetFirstChild<CellValue>();
            cellValue1.Text = "1000";

        }


        // Given a document name and text, 
        // inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
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
