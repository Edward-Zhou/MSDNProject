using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    class RetrieveProperty
    {
        public static void GetPropertyFromDocument(string document)
        {
            XmlDocument xmlProperties = new XmlDocument();

            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(document, false))
            {
                ExtendedFilePropertiesPart appPart = wordDoc.ExtendedFilePropertiesPart;
                xmlProperties.Load(appPart.GetStream());

            }
            XmlNodeList chars = xmlProperties.GetElementsByTagName("Words");

            var xmlNode = chars.Item(0);
            if (xmlNode != null) MessageBox.Show("Word Count:{0}", xmlNode.InnerText);

        }
        public static void MyGetPropertyFromDocument(string document)
        {
            XmlDocument xmlProperties = new XmlDocument();

            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(document, false))
            {
                ExtendedFilePropertiesPart appPart = wordDoc.ExtendedFilePropertiesPart;

                xmlProperties.Load(appPart.GetStream());
            }
            XmlNodeList chars = xmlProperties.GetElementsByTagName("Characters");

            MessageBox.Show("Number of characters in the file = " +
                chars.Item(0).InnerText, "Character Count");
        }

        public static void getAllWords(string document)
        {
            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(document, false))
            {
                IEnumerable<Paragraph> paragraphs = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>();
                int total = 0;
                foreach (Paragraph p in paragraphs)
                {
                    if (p.Descendants<Text>().FirstOrDefault() != null)
                    {
                        string s = p.Descendants<Text>().FirstOrDefault().InnerText;
                        total += s.Split(' ').ToList().Count();
                    }
                    
                }
                MessageBox.Show(total.ToString());

            }
        }
    }
}
