﻿using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System;

namespace OfficeAPI.Class.OpenXmlHelper
{
    public class DateRegion
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
            //Adds new parts or new relationships.
            AddParts();
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

        /// <summary>
        /// Adds new parts or new relationship between parts.
        /// </summary>
        private void AddParts()
        {
            //Generate new parts.
            SharedStringTablePart sharedStringTablePart1 = document.WorkbookPart.AddNewPart<SharedStringTablePart>("rId4");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

        }

        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable();

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2015-09-09T07:03:32Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeWorkbookStylesPart1(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = workbookStylesPart1.Stylesheet;

            Fonts fonts1 = stylesheet1.GetFirstChild<Fonts>();
            CellFormats cellFormats1 = stylesheet1.GetFirstChild<CellFormats>();

            NumberingFormats numberingFormats1 = new NumberingFormats() { Count = (UInt32Value)1U };
            NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)176U, FormatCode = "dd\\.mm\\.yyyy;@" };

            numberingFormats1.Append(numberingFormat1);
            stylesheet1.InsertBefore(numberingFormats1, fonts1);
            cellFormats1.Count = (UInt32Value)2U;

            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)176U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            cellFormats1.Append(cellFormat1);
        }

        private void ChangeWorksheetPart1(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;

            SheetViews sheetViews1 = worksheet1.GetFirstChild<SheetViews>();
            SheetData sheetData1 = worksheet1.GetFirstChild<SheetData>();

            SheetView sheetView1 = sheetViews1.GetFirstChild<SheetView>();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, DyDescent = 0.15D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U };
            CellValue cellValue1 = new CellValue();
            //cellValue1.Text = "42097";
            //cellValue1.Text = DateTime.Now.ToOADate().ToString();
            cellValue1.Text = "04.01.2013";
            cell1.Append(cellValue1);

            row1.Append(cell1);
            sheetData1.Append(row1);
        }

    }
}
