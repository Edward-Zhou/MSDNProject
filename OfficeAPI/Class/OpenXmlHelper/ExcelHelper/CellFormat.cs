using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeAPI.Class.OpenXmlHelper.ExcelHelper
{
    class CellFormat1
    {
        public void cellFormat()
        {
            
            // Create a spreadsheet document by providing a file name.
            string fileName = @"D:\OfficeDev\OpenXML\Excel\H1.xlsx";

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(fileName, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            //var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            //stylesPart.Stylesheet = new Stylesheet();

            //// cell format list
            //stylesPart.Stylesheet.CellFormats = new CellFormats();
            //// empty one for index 0, seems to be required
            //stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            //// cell format references style format 0, font 0, border 0, fill 2 and applies the fill
            //stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
            //stylesPart.Stylesheet.CellFormats.Count = 2;

            //stylesPart.Stylesheet.Save();            
            
            sheets.Append(sheet);
            Worksheet worksheet = new Worksheet();
            SheetData sheetData = new SheetData();
            Row row = new Row();
            Cell cell = new Cell()
            {
                CellReference = "A1",
                DataType = new EnumValue<CellValues>(CellValues.Number),
                CellValue = new CellValue("1000.15")
                
            };

            row.Append(cell);
            sheetData.Append(row);
            worksheet.Append(sheetData);
            worksheetPart.Worksheet = worksheet;
            //if (workbookpart.WorkbookStylesPart == null)
            //{
            //    WorkbookStylesPart stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
            //    stylesPart.Stylesheet = new Stylesheet();
            //    // cell format list
            //    stylesPart.Stylesheet.CellFormats = new CellFormats();
            //    // empty one for index 0, seems to be required
            //    stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            //    // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
            //    stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
            //    stylesPart.Stylesheet.CellFormats.Count = 2;

            //}
            //ChangeWorkbookStylesPart1(workbookpart.WorkbookStylesPart);
            //ChangeWorksheetPart1(worksheetPart);
            // Close the document.
            spreadsheetDocument.Close();

        }
        private void ChangeWorkbookStylesPart1(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = workbookStylesPart1.Stylesheet;

            CellFormats cellFormats1 = stylesheet1.GetFirstChild<CellFormats>();
            cellFormats1.Count = (UInt32Value)2U;

            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)2U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            cellFormats1.Append(cellFormat1);
        }

        private void ChangeWorksheetPart1(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;

            SheetViews sheetViews1 = worksheet1.GetFirstChild<SheetViews>();
            SheetData sheetData1 = worksheet1.GetFirstChild<SheetData>();

            SheetView sheetView1 = sheetViews1.GetFirstChild<SheetView>();

            Selection selection1 = new Selection() { ActiveCell = "H4", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "H4" } };
            sheetView1.Append(selection1);

            Row row1 = sheetData1.GetFirstChild<Row>();

            Cell cell1 = row1.GetFirstChild<Cell>();
            cell1.StyleIndex = (UInt32Value)1U;
        }


    }
}
