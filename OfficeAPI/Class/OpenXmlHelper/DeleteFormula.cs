using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows.Forms;

namespace OfficeAPI.Class.OpenXmlHelper
{
    class DeleteFormula
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
            // Deletes parts only existing in the source package.
            DeleteParts();
            //Changes the contents of the specified parts.
            //ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
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
        /// Deletes parts only existing in the source package.
        /// </summary>
        private void DeleteParts()
        {
            document.WorkbookPart.DeletePart("rId4");
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2015-08-18T02:46:16Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeWorksheetPart1(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;

            SheetData sheetData1 = worksheet1.GetFirstChild<SheetData>();
            
            //Row row1 = sheetData1.GetFirstChild<Row>();
            //Cell cell1 = row1.Elements<Cell>().ElementAt(2);
            //CellFormula cellFormula1 = cell1.GetFirstChild<CellFormula>();
            //cellFormula1.Remove();
            //loop the formula elements and remove them
            foreach (Row row1 in sheetData1.ChildElements)
            {
                Cell cell1 = row1.Elements<Cell>().ElementAt(2);
                
                //MessageBox.Show();
                //CellFormula cellFormula1 = cell1.GetFirstChild<CellFormula>();
                //cellFormula1.Remove();
            }
        }

    }
}
