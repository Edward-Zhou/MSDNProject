using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace OfficeAPI.Class.OpenXmlHelper
{
   public class AddFormula
    {
        private System.Collections.Generic.IDictionary<System.String,OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String,OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String,DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String,DataPart>();
        private SpreadsheetDocument document;

        public void ChangePackage(string filePath)
        {
            using(document = SpreadsheetDocument.Open(filePath, true))
            {
                ChangeParts();
            }
        }

        private  void ChangeParts()
        {
            //Stores the referrences to all the parts in a dictionary.
            BuildUriPartDictionary();
            //Adds new parts or new relationships.
            AddParts();
            //Changes the contents of the specified parts.
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeWorkbookPart1(document.WorkbookPart);
            ChangeWorksheetPart1(((WorksheetPart)UriPartDictionary["/xl/worksheets/sheet1.xml"]));
        }

        /// <summary>
        /// Stores the references to all the parts in the package.
        /// They could be retrieved by their URIs later.
        /// </summary>
        private  void BuildUriPartDictionary()
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
        private  void AddParts()
        {
            //Generate new parts.
            CalculationChainPart calculationChainPart1 = document.WorkbookPart.AddNewPart<CalculationChainPart>("rId4");
            GenerateCalculationChainPart1Content(calculationChainPart1);

        }

        private void GenerateCalculationChainPart1Content(CalculationChainPart calculationChainPart1)
        {
            CalculationChain calculationChain1 = new CalculationChain();
            CalculationCell calculationCell1 = new CalculationCell(){ CellReference = "C1", SheetId = 1, NewLevel = true };

            calculationChain1.Append(calculationCell1);

            calculationChainPart1.CalculationChain = calculationChain1;
        }

        private  void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2015-09-01T05:12:25Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private  void ChangeWorkbookPart1(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = workbookPart1.Workbook;

            CalculationProperties calculationProperties1=workbook1.GetFirstChild<CalculationProperties>();
            calculationProperties1.CalculationId = (UInt32Value)152511U;
        }

        private  void ChangeWorksheetPart1(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;

            SheetViews sheetViews1=worksheet1.GetFirstChild<SheetViews>();
            SheetData sheetData1=worksheet1.GetFirstChild<SheetData>();

            SheetView sheetView1=sheetViews1.GetFirstChild<SheetView>();

            Selection selection1=sheetView1.GetFirstChild<Selection>();
            selection1.ActiveCell = "E3";
            selection1.SequenceOfReferences = new ListValue<StringValue>() { InnerText = "E3" };

            Row row1=sheetData1.GetFirstChild<Row>();

            Cell cell1=row1.Elements<Cell>().ElementAt(2);

            CellValue cellValue1=cell1.GetFirstChild<CellValue>();

            CellFormula cellFormula1 = new CellFormula();
            cellFormula1.Text = "A1+B1";
            cell1.InsertBefore(cellFormula1,cellValue1);
        }
    }
}
