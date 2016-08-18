using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    class DeleteCommandbar
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
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeMainDocumentPart1(document.MainDocumentPart);
            ChangeWebSettingsPart1(((WebSettingsPart)UriPartDictionary["/word/webSettings.xml"]));
            ChangeDocumentSettingsPart1(((DocumentSettingsPart)UriPartDictionary["/word/settings.xml"]));
            ChangeStyleDefinitionsPart1(((StyleDefinitionsPart)UriPartDictionary["/word/styles.xml"]));
            ChangeGlossaryDocumentPart1(((GlossaryDocumentPart)UriPartDictionary["/word/glossary/document.xml"]));
            ChangeDocumentSettingsPart2(((DocumentSettingsPart)UriPartDictionary["/word/glossary/settings.xml"]));
            ChangeStyleDefinitionsPart2(((StyleDefinitionsPart)UriPartDictionary["/word/glossary/styles.xml"]));
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
            UriPartDictionary["/word/customizations.xml"].DeletePart("rId1");
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

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Creator = "LabAdmin";
            package.PackageProperties.Revision = "1";
            package.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2016-03-17T05:40:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2016-03-17T05:40:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeMainDocumentPart1(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = mainDocumentPart1.Document;

            Body body1 = document1.GetFirstChild<Body>();

            Paragraph paragraph1 = body1.GetFirstChild<Paragraph>();
            SectionProperties sectionProperties1 = body1.GetFirstChild<SectionProperties>();
            paragraph1.RsidParagraphAddition = "00000000";
            paragraph1.RsidRunAdditionDefault = "00000000";
            sectionProperties1.RsidR = "00000000";
            sectionProperties1.RsidSect = null;
        }

        private void ChangeWebSettingsPart1(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = webSettingsPart1.WebSettings;

            TargetScreenSize targetScreenSize1 = new TargetScreenSize() { Val = TargetScreenSizeValues.Sz544x376 };
            webSettings1.Append(targetScreenSize1);
        }

        private void ChangeDocumentSettingsPart1(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = documentSettingsPart1.Settings;

            Rsids rsids1 = settings1.GetFirstChild<Rsids>();

            rsids1.Remove();
        }

        private void ChangeStyleDefinitionsPart1(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = styleDefinitionsPart1.Styles;

            Style style1 = styles1.GetFirstChild<Style>();

            Rsid rsid1 = style1.GetFirstChild<Rsid>();

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
            DocPart docPart5 = docParts1.Elements<DocPart>().ElementAt(4);
            DocPart docPart6 = docParts1.Elements<DocPart>().ElementAt(5);
            DocPart docPart7 = docParts1.Elements<DocPart>().ElementAt(6);
            DocPart docPart8 = docParts1.Elements<DocPart>().ElementAt(7);
            DocPart docPart9 = docParts1.Elements<DocPart>().ElementAt(8);
            DocPart docPart10 = docParts1.Elements<DocPart>().ElementAt(9);
            DocPart docPart11 = docParts1.Elements<DocPart>().ElementAt(10);
            DocPart docPart12 = docParts1.Elements<DocPart>().ElementAt(11);
            DocPart docPart13 = docParts1.Elements<DocPart>().ElementAt(12);
            DocPart docPart14 = docParts1.Elements<DocPart>().ElementAt(13);

            DocPartProperties docPartProperties1 = docPart1.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody1 = docPart1.GetFirstChild<DocPartBody>();

            DocPartId docPartId1 = docPartProperties1.GetFirstChild<DocPartId>();
            docPartId1.Val = "{DB6F3F57-D044-490D-862E-CDBA69AFA6B5}";

            Paragraph paragraph1 = docPartBody1.GetFirstChild<Paragraph>();
            paragraph1.RsidParagraphAddition = "00000000";
            paragraph1.RsidParagraphProperties = "00E51E4E";
            paragraph1.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties2 = docPart2.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody2 = docPart2.GetFirstChild<DocPartBody>();

            DocPartId docPartId2 = docPartProperties2.GetFirstChild<DocPartId>();
            docPartId2.Val = "{D1BB5457-8E5E-46FA-B5F1-94CF3E2E0CD2}";

            Paragraph paragraph2 = docPartBody2.GetFirstChild<Paragraph>();
            paragraph2.RsidParagraphAddition = "00000000";
            paragraph2.RsidParagraphProperties = "00E51E4E";
            paragraph2.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties3 = docPart3.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody3 = docPart3.GetFirstChild<DocPartBody>();

            DocPartId docPartId3 = docPartProperties3.GetFirstChild<DocPartId>();
            docPartId3.Val = "{49A6CA21-C21D-4AE7-8426-093D0A7D8ADD}";

            Paragraph paragraph3 = docPartBody3.GetFirstChild<Paragraph>();
            paragraph3.RsidParagraphAddition = "00000000";
            paragraph3.RsidParagraphProperties = "00E51E4E";
            paragraph3.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties4 = docPart4.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody4 = docPart4.GetFirstChild<DocPartBody>();

            DocPartId docPartId4 = docPartProperties4.GetFirstChild<DocPartId>();
            docPartId4.Val = "{66070641-D14E-44D7-A456-85E6136A2431}";

            Paragraph paragraph4 = docPartBody4.GetFirstChild<Paragraph>();
            paragraph4.RsidParagraphAddition = "00000000";
            paragraph4.RsidParagraphProperties = "00E51E4E";
            paragraph4.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties5 = docPart5.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody5 = docPart5.GetFirstChild<DocPartBody>();

            DocPartId docPartId5 = docPartProperties5.GetFirstChild<DocPartId>();
            docPartId5.Val = "{B3D65FDF-951C-4AD4-B125-7F65D2C83AF1}";

            Paragraph paragraph5 = docPartBody5.GetFirstChild<Paragraph>();
            paragraph5.RsidParagraphAddition = "00000000";
            paragraph5.RsidParagraphProperties = "00E51E4E";
            paragraph5.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties6 = docPart6.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody6 = docPart6.GetFirstChild<DocPartBody>();

            DocPartId docPartId6 = docPartProperties6.GetFirstChild<DocPartId>();
            docPartId6.Val = "{3579B2FB-5997-40FD-BD12-1C07B8A5E1A2}";

            Paragraph paragraph6 = docPartBody6.GetFirstChild<Paragraph>();
            paragraph6.RsidParagraphAddition = "00000000";
            paragraph6.RsidParagraphProperties = "00E51E4E";
            paragraph6.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties7 = docPart7.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody7 = docPart7.GetFirstChild<DocPartBody>();

            DocPartId docPartId7 = docPartProperties7.GetFirstChild<DocPartId>();
            docPartId7.Val = "{528EE693-D4AA-401E-8B61-44930D363AAB}";

            Paragraph paragraph7 = docPartBody7.GetFirstChild<Paragraph>();
            paragraph7.RsidParagraphAddition = "00000000";
            paragraph7.RsidParagraphProperties = "00E51E4E";
            paragraph7.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties8 = docPart8.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody8 = docPart8.GetFirstChild<DocPartBody>();

            DocPartId docPartId8 = docPartProperties8.GetFirstChild<DocPartId>();
            docPartId8.Val = "{2B4FA776-C743-4F5C-8F50-31E8C938997F}";

            Paragraph paragraph8 = docPartBody8.GetFirstChild<Paragraph>();
            paragraph8.RsidParagraphAddition = "00000000";
            paragraph8.RsidParagraphProperties = "00E51E4E";
            paragraph8.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties9 = docPart9.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody9 = docPart9.GetFirstChild<DocPartBody>();

            DocPartId docPartId9 = docPartProperties9.GetFirstChild<DocPartId>();
            docPartId9.Val = "{828EFC50-614C-46DF-8826-BD9D81186AB7}";

            Paragraph paragraph9 = docPartBody9.GetFirstChild<Paragraph>();
            paragraph9.RsidParagraphAddition = "00000000";
            paragraph9.RsidParagraphProperties = "00E51E4E";
            paragraph9.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties10 = docPart10.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody10 = docPart10.GetFirstChild<DocPartBody>();

            DocPartId docPartId10 = docPartProperties10.GetFirstChild<DocPartId>();
            docPartId10.Val = "{874CA0C5-284B-442C-BDB1-67B75F150727}";

            Paragraph paragraph10 = docPartBody10.GetFirstChild<Paragraph>();
            paragraph10.RsidParagraphAddition = "00000000";
            paragraph10.RsidParagraphProperties = "00E51E4E";
            paragraph10.RsidRunAdditionDefault = "00E51E4E";

            DocPartProperties docPartProperties11 = docPart11.GetFirstChild<DocPartProperties>();
            DocPartBody docPartBody11 = docPart11.GetFirstChild<DocPartBody>();

            DocPartId docPartId11 = docPartProperties11.GetFirstChild<DocPartId>();
            docPartId11.Val = "{A7EA2502-BCF8-402F-BF0B-D63960CAF4C5}";

            Paragraph paragraph11 = docPartBody11.GetFirstChild<Paragraph>();
            paragraph11.RsidParagraphAddition = "00000000";
            paragraph11.RsidParagraphProperties = "00E51E4E";
            paragraph11.RsidRunAdditionDefault = "00E51E4E";

            docPart12.Remove();
            docPart13.Remove();
            docPart14.Remove();
        }

        private void ChangeDocumentSettingsPart2(DocumentSettingsPart documentSettingsPart2)
        {
            Settings settings1 = documentSettingsPart2.Settings;

            Rsids rsids1 = settings1.GetFirstChild<Rsids>();

            RsidRoot rsidRoot1 = rsids1.GetFirstChild<RsidRoot>();
            Rsid rsid1 = rsids1.GetFirstChild<Rsid>();
            Rsid rsid2 = rsids1.Elements<Rsid>().ElementAt(1);
            rsidRoot1.Val = "00E51E4E";
            rsid1.Val = "00E51E4E";

            rsid2.Remove();
        }

        private void ChangeStyleDefinitionsPart2(StyleDefinitionsPart styleDefinitionsPart2)
        {
            Styles styles1 = styleDefinitionsPart2.Styles;

            Style style1 = styles1.GetFirstChild<Style>();
            Style style2 = styles1.Elements<Style>().ElementAt(4);

            Rsid rsid1 = style1.GetFirstChild<Rsid>();

            rsid1.Remove();

            Rsid rsid2 = style2.GetFirstChild<Rsid>();
            rsid2.Val = "00E51E4E";
        }

    }
}
