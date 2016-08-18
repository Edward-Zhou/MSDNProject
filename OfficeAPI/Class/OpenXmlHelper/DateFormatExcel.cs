using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Windows.Forms;

namespace OfficeAPI.Class.OpenXmlHelper
{
    public class DateFormatExcel
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private SpreadsheetDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = SpreadsheetDocument.Open(filePath, true))
            {
                ChangeParts();
            }
        }

        private void ChangeParts()
        {
            //Stores the referrences to all the parts in a dictionary.
            BuildUriPartDictionary();
            //Changes the contents of the specified parts.
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeWorkbookStylesPart1(((WorkbookStylesPart)UriPartDictionary["/xl/styles.xml"]));
            ChangeWorksheetPart1(((WorksheetPart)UriPartDictionary["/xl/worksheets/sheet1.xml"]));
        }

        /// <summary>
        /// Stores the references to all the parts in the package.
        /// They could be retrieved by their URIs later.
        /// </summary>
        private void BuildUriPartDictionary()
        {
            System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
            queue.Enqueue(document);
            while (queue.Count > 0)
            {
                foreach (var part in queue.Dequeue().Parts)
                {
                    if (!UriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString()))
                    {
                        UriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                        queue.Enqueue(part.OpenXmlPart);
                    }
                }
            }
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2015-09-07T01:48:33Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeWorkbookStylesPart1(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = workbookStylesPart1.Stylesheet;

            CellFormats cellFormats1 = stylesheet1.GetFirstChild<CellFormats>();
            cellFormats1.Count = (UInt32Value)2U;

            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            cellFormats1.Append(cellFormat1);
        }

        private void ChangeWorksheetPart1(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;

            SheetViews sheetViews1 = worksheet1.GetFirstChild<SheetViews>();
            SheetData sheetData1 = worksheet1.GetFirstChild<SheetData>();

            SheetView sheetView1 = sheetViews1.GetFirstChild<SheetView>();

            Selection selection1 = sheetView1.GetFirstChild<Selection>();
            selection1.ActiveCell = "F12";
            selection1.SequenceOfReferences = new ListValue<StringValue>() { InnerText = "F12" };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 9.5D, BestFit = true, CustomWidth = true };

            columns1.Append(column1);
            worksheet1.InsertBefore(columns1, sheetData1);

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, DyDescent = 0.15D };
            //format the cell style with style index
            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U };

            Cell cell2 = GetCell(worksheetPart1.Worksheet, "A",2);
            Cell cell3 = GetCell(worksheetPart1.Worksheet, "A", 3);
            Cell cell5 = GetCell(worksheetPart1.Worksheet, "A", 5);
            Cell cell6 = GetCell(worksheetPart1.Worksheet, "B", 5);
            MessageBox.Show(cell6.LastChild.ToString());
            cell2.StyleIndex = cell5.StyleIndex;
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "Hello World Hello wORD";
            cell2.Append(cellValue2);
            CellValue cellValue1 = new CellValue();
            DateTime dt=DateTime.Now;
            //cellValue1.Text = "42254";
            cellValue1.Text = dt.ToOADate().ToString();

            cell1.Append(cellValue1);

            row1.Append(cell1);
            sheetData1.Append(row1);
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
