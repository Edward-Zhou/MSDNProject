using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    class FormProtect
    {
        private WordprocessingDocument document;
        public void formProtect(string filePath)
        {
            using (document = WordprocessingDocument.Open(filePath, true))
            {
                MainDocumentPart mainDocumentPart1 = document.MainDocumentPart;
            }
        }
        public void ChangeMainDocumentPart(MainDocumentPart mainDocumentPart1)
        {

            Document document1 = mainDocumentPart1.Document;

            Body body1 = document1.GetFirstChild<Body>();

            Paragraph paragraph1 = body1.GetFirstChild<Paragraph>();
            Paragraph paragraph2 = body1.Elements<Paragraph>().ElementAt(1);
            Paragraph paragraph3 = body1.Elements<Paragraph>().ElementAt(2);
            Paragraph paragraph4 = body1.Elements<Paragraph>().ElementAt(3);
            Paragraph paragraph5 = body1.Elements<Paragraph>().ElementAt(4);
            Paragraph paragraph6 = body1.Elements<Paragraph>().ElementAt(6);

            Run run1 = paragraph1.GetFirstChild<Run>();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            paragraphMarkRunProperties1.Append(runFonts1);

            paragraphProperties1.Append(paragraphMarkRunProperties1);
            paragraph1.InsertBefore(paragraphProperties1, run1);

            SdtRun sdtRun1 = paragraph2.GetFirstChild<SdtRun>();

            SdtEndCharProperties sdtEndCharProperties1 = sdtRun1.GetFirstChild<SdtEndCharProperties>();

            sdtEndCharProperties1.Remove();

            SdtRun sdtRun2 = paragraph3.GetFirstChild<SdtRun>();

            SdtEndCharProperties sdtEndCharProperties2 = sdtRun2.GetFirstChild<SdtEndCharProperties>();

            sdtEndCharProperties2.Remove();

            SdtRun sdtRun3 = paragraph4.GetFirstChild<SdtRun>();

            SdtEndCharProperties sdtEndCharProperties3 = sdtRun3.GetFirstChild<SdtEndCharProperties>();

            sdtEndCharProperties3.Remove();

            SdtRun sdtRun4 = paragraph5.GetFirstChild<SdtRun>();

            SdtProperties sdtProperties1 = sdtRun4.GetFirstChild<SdtProperties>();
            SdtEndCharProperties sdtEndCharProperties4 = sdtRun4.GetFirstChild<SdtEndCharProperties>();
            SdtContentRun sdtContentRun1 = sdtRun4.GetFirstChild<SdtContentRun>();

            SdtPlaceholder sdtPlaceholder1 = sdtProperties1.GetFirstChild<SdtPlaceholder>();

            Lock lock1 = new Lock() { Val = LockingValues.SdtContentLocked };
            sdtProperties1.InsertBefore(lock1, sdtPlaceholder1);

            sdtEndCharProperties4.Remove();

            Run run2 = sdtContentRun1.GetFirstChild<Run>();

            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            sdtContentRun1.InsertBefore(proofError1, run2);

            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            sdtContentRun1.Append(proofError2);

            Run run3 = paragraph6.GetFirstChild<Run>();

            Text text1 = run3.GetFirstChild<Text>();
            text1.Text = "Hello Wor";


            Run run4 = new Run();
            Text text2 = new Text();
            text2.Text = "d";

            run4.Append(text2);
            paragraph6.Append(run4);
        }

    }
}
