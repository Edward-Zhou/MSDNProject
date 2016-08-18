using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    class FormProTotal
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private WordprocessingDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = WordprocessingDocument.Open(filePath, true))
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
            //Changes the relationship ID of the parts.
            ReconfigureRelationshipID();
            //Changes the contents of the specified parts.
            ChangeExtendedFilePropertiesPart1(((ExtendedFilePropertiesPart)UriPartDictionary["/docProps/app.xml"]));
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeMainDocumentPart1(document.MainDocumentPart);
            ChangeDocumentSettingsPart1(((DocumentSettingsPart)UriPartDictionary["/word/settings.xml"]));
            ChangeGlossaryDocumentPart1(((GlossaryDocumentPart)UriPartDictionary["/word/glossary/document.xml"]));
            ChangeDocumentSettingsPart2(((DocumentSettingsPart)UriPartDictionary["/word/glossary/settings.xml"]));
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
            UriPartDictionary["/customXml/item1.xml"].DeletePart("rId1");
            document.MainDocumentPart.DeletePart("rId1");
        }

        /// <summary>
        /// Changes the relationship ID of the parts in the source package to make sure these IDs are the same as those in the target package.
        /// To avoid the conflict of the relationship ID, a temporary ID is assigned first.        
        /// </summary>
        private void ReconfigureRelationshipID()
        {
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/settings.xml"], "generatedTmpID1");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/theme/theme1.xml"], "generatedTmpID2");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/styles.xml"], "generatedTmpID3");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/glossary/document.xml"], "generatedTmpID4");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/fontTable.xml"], "generatedTmpID5");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/webSettings.xml"], "generatedTmpID6");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/settings.xml"], "rId2");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/theme/theme1.xml"], "rId6");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/styles.xml"], "rId1");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/glossary/document.xml"], "rId5");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/fontTable.xml"], "rId4");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/webSettings.xml"], "rId3");
        }

        private void ChangeExtendedFilePropertiesPart1(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = extendedFilePropertiesPart1.Properties;

            Ap.TotalTime totalTime1 = properties1.GetFirstChild<Ap.TotalTime>();
            totalTime1.Text = "9";

        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Revision = "3";
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2015-09-18T05:40:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeMainDocumentPart1(MainDocumentPart mainDocumentPart1)
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

        private void ChangeDocumentSettingsPart1(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = documentSettingsPart1.Settings;

            FormsDesign formsDesign1 = settings1.GetFirstChild<FormsDesign>();
            Rsids rsids1 = settings1.GetFirstChild<Rsids>();

            formsDesign1.Remove();

            Rsid rsid1 = rsids1.Elements<Rsid>().ElementAt(2);

            rsid1.Remove();
        }

        private void ChangeGlossaryDocumentPart1(GlossaryDocumentPart glossaryDocumentPart1)
        {
            GlossaryDocument glossaryDocument1 = glossaryDocumentPart1.GlossaryDocument;

            DocParts docParts1 = glossaryDocument1.GetFirstChild<DocParts>();

            DocPart docPart1 = docParts1.GetFirstChild<DocPart>();
            DocPart docPart2 = docParts1.Elements<DocPart>().ElementAt(1);
            DocPart docPart3 = docParts1.Elements<DocPart>().ElementAt(2);
            DocPart docPart4 = docParts1.Elements<DocPart>().ElementAt(3);

            DocPartBody docPartBody1 = docPart1.GetFirstChild<DocPartBody>();

            Paragraph paragraph1 = docPartBody1.GetFirstChild<Paragraph>();
            paragraph1.RsidParagraphAddition = "00000000";

            DocPartBody docPartBody2 = docPart2.GetFirstChild<DocPartBody>();

            Paragraph paragraph2 = docPartBody2.GetFirstChild<Paragraph>();
            paragraph2.RsidParagraphAddition = "00000000";

            DocPartBody docPartBody3 = docPart3.GetFirstChild<DocPartBody>();

            Paragraph paragraph3 = docPartBody3.GetFirstChild<Paragraph>();
            paragraph3.RsidParagraphAddition = "00000000";

            DocPartBody docPartBody4 = docPart4.GetFirstChild<DocPartBody>();

            Paragraph paragraph4 = docPartBody4.GetFirstChild<Paragraph>();
            paragraph4.RsidParagraphAddition = "00000000";
        }

        private void ChangeDocumentSettingsPart2(DocumentSettingsPart documentSettingsPart2)
        {
            Settings settings1 = documentSettingsPart2.Settings;

            Rsids rsids1 = settings1.GetFirstChild<Rsids>();

            Rsid rsid1 = rsids1.Elements<Rsid>().ElementAt(1);
            Rsid rsid2 = rsids1.Elements<Rsid>().ElementAt(2);

            rsid1.Remove();
            rsid2.Remove();
        }

    }
}
