using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeAPI.Class.OpenXmlHelper
{
    class AcceptRevisionsClass
    {
        public static void AcceptRevisions(string fileName, string authorName)
        {
            // Given a document name and an author name, accept revisions. 
            using (WordprocessingDocument wdDoc =
                WordprocessingDocument.Open(fileName, true))
            {
                Body body = wdDoc.MainDocumentPart.Document.Body;
                // Handle the table formatting changes.
                List<OpenXmlElement> tablechanges =
                    body.Descendants<TablePropertiesChange>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

                foreach (OpenXmlElement change in tablechanges)
                {
                    change.Remove();
                }
                // Handle the table cell formatting changes.
                List<OpenXmlElement> tablecellchanges =
                    body.Descendants<TableCellPropertiesChange>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

                foreach (OpenXmlElement change in tablechanges)
                {
                    change.Remove();
                }
                // Handle the formatting changes.
                List<OpenXmlElement> changes =
                    body.Descendants<ParagraphPropertiesChange>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

                foreach (OpenXmlElement change in changes)
                {
                    change.Remove();
                }

                // Handle the deletions.
                List<OpenXmlElement> deletions =
                    body.Descendants<Deleted>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

                deletions.AddRange(body.Descendants<DeletedRun>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

                deletions.AddRange(body.Descendants<DeletedMathControl>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

                foreach (OpenXmlElement deletion in deletions)
                {
                    deletion.Remove();
                }

                // Handle the insertions.
                List<OpenXmlElement> insertions =
                    body.Descendants<Inserted>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

                insertions.AddRange(body.Descendants<InsertedRun>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

                insertions.AddRange(body.Descendants<InsertedMathControl>()
                    .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

                foreach (OpenXmlElement insertion in insertions)
                {
                    // Found new content.
                    // Promote them to the same level as node, and then delete the node.
                    foreach (var run in insertion.Elements<Run>())
                    {
                        if (run == insertion.FirstChild)
                        {
                            insertion.InsertAfterSelf(new Run(run.OuterXml));
                        }
                        else
                        {
                            insertion.NextSibling().InsertAfterSelf(new Run(run.OuterXml));
                        }
                    }
                    insertion.RemoveAttribute("rsidR",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    insertion.RemoveAttribute("rsidRPr",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    insertion.Remove();
                }
            }
        }
        public static System.Type[] trackedRevisionsElements = new System.Type[] {
        //typeof(CellDeletion),
        //typeof(CellInsertion),
        //typeof(CellMerge),
        //typeof(CustomXmlDelRangeEnd),
        //typeof(CustomXmlDelRangeStart),
        //typeof(CustomXmlInsRangeEnd),
        //typeof(CustomXmlInsRangeStart),
        //typeof(Deleted),
        //typeof(DeletedFieldCode),
        //typeof(DeletedMathControl),
        //typeof(DeletedRun),
        //typeof(DeletedText),
        //typeof(Inserted),
        //typeof(InsertedMathControl),
        //typeof(InsertedMathControl),
        //typeof(InsertedRun),
        //typeof(MoveFrom),
        //typeof(MoveFromRangeEnd),
        //typeof(MoveFromRangeStart),
        //typeof(MoveTo),
        //typeof(MoveToRangeEnd),
        //typeof(MoveToRangeStart),
        typeof(MoveToRun),
        typeof(NumberingChange),
        typeof(ParagraphMarkRunPropertiesChange),
        typeof(ParagraphPropertiesChange),
        typeof(RunPropertiesChange),
        typeof(SectionPropertiesChange),
        //typeof(TableCellPropertiesChange),
        //typeof(TableGridChange),
        //typeof(TablePropertiesChange),
        //typeof(TablePropertyExceptionsChange),
        //typeof(TableRowPropertiesChange),
    };

        public static bool PartHasTrackedRevisions(OpenXmlPart part)
        {
            bool b = part.RootElement.Descendants()
                .Any(e => trackedRevisionsElements.Contains(e.GetType()));
            //if (b)
            //{
            //    MessageBox.Show(part.RootElement.Descendants().ToList().ToString());
            //}
            return b;
        }

        public static bool HasTrackedRevisions(WordprocessingDocument doc)
        {
            if (PartHasTrackedRevisions(doc.MainDocumentPart))
                return true;
            foreach (var part in doc.MainDocumentPart.HeaderParts)
                if (PartHasTrackedRevisions(part))
                    return true;
            foreach (var part in doc.MainDocumentPart.FooterParts)
                if (PartHasTrackedRevisions(part))
                    return true;
            if (doc.MainDocumentPart.EndnotesPart != null)
                if (PartHasTrackedRevisions(doc.MainDocumentPart.EndnotesPart))
                    return true;
            if (doc.MainDocumentPart.FootnotesPart != null)
                if (PartHasTrackedRevisions(doc.MainDocumentPart.FootnotesPart))
                    return true;
            return false;
        }

    }
}
