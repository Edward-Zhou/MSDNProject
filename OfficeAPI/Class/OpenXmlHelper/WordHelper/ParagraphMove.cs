using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;


namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    public class ParagraphMove
    {
        private WordprocessingDocument document;
        public void paragraphMove(string filePath)
        {
            using (document = WordprocessingDocument.Open(filePath, true))
            {
                Paragraph paragraph1 = (Paragraph)document.MainDocumentPart.Document.Body.Descendants<Paragraph>().ElementAt(1);
                Paragraph paragraph3= (Paragraph)document.MainDocumentPart.Document.Body.Descendants<Paragraph>().ElementAt(3);
                Paragraph paragraph4 = (Paragraph)document.MainDocumentPart.Document.Body.Descendants<Paragraph>().ElementAt(4);
                Table table = (Table)document.MainDocumentPart.Document.Body.Descendants<Table>().FirstOrDefault();

                Body body = document.MainDocumentPart.Document.GetFirstChild<Body>();
                Paragraph p = new Paragraph();
                paragraph4.Append(p);
                paragraph4.Append(paragraph1.CloneNode(true));
                paragraph4.Append(table.CloneNode(true));
                paragraph4.Append(paragraph3.CloneNode(true));
                paragraph1.Remove();
                table.Remove();
                paragraph3.Remove();
            }
        }
    }
}
