using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace GeneratedCode
{
    public class AddString
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
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)1U, UniqueCount = (UInt32Value)1U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text();
            text1.Text = "String";
            PhoneticProperties phoneticProperties1 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem1.Append(text1);
            sharedStringItem1.Append(phoneticProperties1);

            sharedStringTable1.Append(sharedStringItem1);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2016-06-03T06:03:20Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeWorksheetPart1(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;

            SheetDimension sheetDimension1 = worksheet1.GetFirstChild<SheetDimension>();
            SheetViews sheetViews1 = worksheet1.GetFirstChild<SheetViews>();
            SheetData sheetData1 = worksheet1.GetFirstChild<SheetData>();
            sheetDimension1.Reference = "A1:AA6";

            SheetView sheetView1 = sheetViews1.GetFirstChild<SheetView>();

            Selection selection1 = new Selection() { ActiveCell = "AA7", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "AA7" } };
            sheetView1.Append(selection1);

            Row row1 = sheetData1.GetFirstChild<Row>();
            row1.Spans = new ListValue<StringValue>() { InnerText = "1:27" };

            Row row2 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell1 = new Cell() { CellReference = "AA6", DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "0";

            cell1.Append(cellValue1);

            row2.Append(cell1);
            sheetData1.Append(row2);
        }


    }
}
