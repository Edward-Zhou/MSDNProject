using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace OfficeAPI.Class.OpenXmlHelper
{
    class GetTableCountClass
    {
        public static void getTableCount()
        {
            using (WordprocessingDocument package = WordprocessingDocument.Open(@"D:\OfficeDev\OpenXML\TableCount.docx", false))
            {
                //var tables = mainPart.Document.Descendants<Table>().ToList();
                List<Table> tables = package.MainDocumentPart.Document.Descendants<Table>().ToList();
                MessageBox.Show(tables.Count().ToString());
                List<Table> tables1 = package.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                MessageBox.Show(tables1.Count().ToString());
                //List<Table> table2 = package.MainDocumentPart.HeaderParts.FirstOrDefault().g
                OpenXmlPart part = package.MainDocumentPart.HeaderParts.FirstOrDefault();
                
                //part.Parts.
                XDocument xd = part.GetXDocument();
                //xd.Elements("w:tbl").Count();
                MessageBox.Show(xd.Elements().Count().ToString());
                //List<Table> table2=xd.Descendants()

            }
        }
    }
    public static class Extensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
                return partXDocument;
            using (Stream partStream = part.GetStream())
            using (XmlReader partXmlReader = XmlReader.Create(partStream))
                partXDocument = XDocument.Load(partXmlReader);
            part.AddAnnotation(partXDocument);
            return partXDocument;
        }
    }

}
