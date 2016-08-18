using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    class Orientations
    {
        WordprocessingDocument package;

        public void ChangeOrientation(string filePath)
        {
            using (package = WordprocessingDocument.Open(filePath, true))
            {
                var sp = package.MainDocumentPart.Document.Body.Descendants<SectionProperties>();
                foreach (SectionProperties sectPr in sp)
                {
                    PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
                    pgSz.Orient = new EnumValue<PageOrientationValues>(PageOrientationValues.Landscape);

                }
                package.MainDocumentPart.Document.Save();
            }
        }
        public void ChangeOrientationbak(string filePath)
        {
            using (package = WordprocessingDocument.Open(filePath, true))
            {                
                Body body = package.MainDocumentPart.Document.Body;
                Paragraph[] para = body.Descendants<Paragraph>().ToArray();  //.ToArray();

                for (int i = 0; i < para.Count(); i++)
                {
                    ParagraphProperties paragraphProperties1 = para[i].GetFirstChild<ParagraphProperties>();
                    if (i%2 == 0) //odd page
                    {                       

                        SectionProperties sectionProperties1 = paragraphProperties1.GetFirstChild<SectionProperties>();
                        sectionProperties1.RsidSect = "00BA2F0F";

                        PageSize pageSize1 = sectionProperties1.GetFirstChild<PageSize>();

                        PageSize pageSize2 = new PageSize() { Width = (UInt32Value)15840U, Height = (UInt32Value)12240U, Orient = PageOrientationValues.Landscape };
                        sectionProperties1.InsertBefore(pageSize2, pageSize1);

                        pageSize1.Remove();

                        //CreateBreakBefore();
                        //ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                        //SectionProperties sectionProperties2 = new SectionProperties() { RsidR = "00BA2F0F", RsidSect = "00BA2F0F" };
                        //PageSize pageSize2 = new PageSize() { Width = (UInt32Value)15840U, Height = (UInt32Value)12240U, Orient = PageOrientationValues.Landscape };
                        //PageMargin pageMargin2 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
                        //Columns columns2 = new Columns() { Space = "720" };
                        //DocGrid docGrid2 = new DocGrid() { LinePitch = 360 };
                        //sectionProperties2.Append(pageSize2);
                        //sectionProperties2.Append(pageMargin2);
                        //sectionProperties2.Append(columns2);
                        //sectionProperties2.Append(docGrid2);
                        //paragraphProperties2.Append(sectionProperties2);
                        //para[i].Append(paragraphProperties2);
                        //CreateBreakAfter();
                    }
                }
                
                //Paragraph para1 = CreatePara1(); // Paragraph 1 Portrait

                //Paragraph breakBefore = CreateBreakBefore(); // Break before next paragraph that would be landscape

                //Paragraph para2 = CreatePara2(); // Paragraph 2 Landscape

                //Paragraph breakAfter = CreateBreakAfter(); // Break after previous paragraph that was landscape

                //Paragraph para3 = CreatePara3(); // Paragraph 3 Portrait again

                //body.Append(para1);
                //body.Append(breakBefore);
                //body.Append(para2);
                //body.Append(breakAfter);
                //body.Append(para3);

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
            Paragraph paragraph6 = body1.Elements<Paragraph>().ElementAt(5);
            Paragraph paragraph7 = body1.Elements<Paragraph>().ElementAt(6);
            Paragraph paragraph8 = body1.Elements<Paragraph>().ElementAt(7);
            Paragraph paragraph9 = body1.Elements<Paragraph>().ElementAt(8);
            Paragraph paragraph10 = body1.Elements<Paragraph>().ElementAt(9);
            Paragraph paragraph11 = body1.Elements<Paragraph>().ElementAt(10);
            Paragraph paragraph12 = body1.Elements<Paragraph>().ElementAt(11);
            Paragraph paragraph13 = body1.Elements<Paragraph>().ElementAt(12);
            Paragraph paragraph14 = body1.Elements<Paragraph>().ElementAt(13);
            Paragraph paragraph15 = body1.Elements<Paragraph>().ElementAt(14);
            Paragraph paragraph16 = body1.Elements<Paragraph>().ElementAt(15);
            Paragraph paragraph17 = body1.Elements<Paragraph>().ElementAt(16);
            Paragraph paragraph18 = body1.Elements<Paragraph>().ElementAt(17);
            Paragraph paragraph19 = body1.Elements<Paragraph>().ElementAt(18);
            Paragraph paragraph20 = body1.Elements<Paragraph>().ElementAt(19);
            Paragraph paragraph21 = body1.Elements<Paragraph>().ElementAt(20);
            Paragraph paragraph22 = body1.Elements<Paragraph>().ElementAt(21);
            Paragraph paragraph23 = body1.Elements<Paragraph>().ElementAt(22);
            Paragraph paragraph24 = body1.Elements<Paragraph>().ElementAt(23);
            Paragraph paragraph25 = body1.Elements<Paragraph>().ElementAt(24);
            Paragraph paragraph26 = body1.Elements<Paragraph>().ElementAt(25);
            Paragraph paragraph27 = body1.Elements<Paragraph>().ElementAt(26);
            Paragraph paragraph28 = body1.Elements<Paragraph>().ElementAt(27);
            Paragraph paragraph29 = body1.Elements<Paragraph>().ElementAt(28);
            Paragraph paragraph30 = body1.Elements<Paragraph>().ElementAt(29);
            Paragraph paragraph31 = body1.Elements<Paragraph>().ElementAt(30);
            Paragraph paragraph32 = body1.Elements<Paragraph>().ElementAt(31);
            Paragraph paragraph33 = body1.Elements<Paragraph>().ElementAt(32);
            Paragraph paragraph34 = body1.Elements<Paragraph>().ElementAt(33);
            Paragraph paragraph35 = body1.Elements<Paragraph>().ElementAt(34);
            Paragraph paragraph36 = body1.Elements<Paragraph>().ElementAt(35);
            Paragraph paragraph37 = body1.Elements<Paragraph>().ElementAt(36);
            Paragraph paragraph38 = body1.Elements<Paragraph>().ElementAt(37);
            Paragraph paragraph39 = body1.Elements<Paragraph>().ElementAt(38);
            Paragraph paragraph40 = body1.Elements<Paragraph>().ElementAt(39);
            Paragraph paragraph41 = body1.Elements<Paragraph>().ElementAt(40);
            Paragraph paragraph42 = body1.Elements<Paragraph>().ElementAt(41);
            Paragraph paragraph43 = body1.Elements<Paragraph>().ElementAt(42);
            Paragraph paragraph44 = body1.Elements<Paragraph>().ElementAt(43);
            Paragraph paragraph45 = body1.Elements<Paragraph>().ElementAt(44);
            Paragraph paragraph46 = body1.Elements<Paragraph>().ElementAt(45);
            Paragraph paragraph47 = body1.Elements<Paragraph>().ElementAt(46);
            Paragraph paragraph48 = body1.Elements<Paragraph>().ElementAt(47);
            Paragraph paragraph49 = body1.Elements<Paragraph>().ElementAt(48);
            Paragraph paragraph50 = body1.Elements<Paragraph>().ElementAt(49);
            Paragraph paragraph51 = body1.Elements<Paragraph>().ElementAt(50);
            Paragraph paragraph52 = body1.Elements<Paragraph>().ElementAt(51);
            Paragraph paragraph53 = body1.Elements<Paragraph>().ElementAt(52);
            Paragraph paragraph54 = body1.Elements<Paragraph>().ElementAt(53);
            Paragraph paragraph55 = body1.Elements<Paragraph>().ElementAt(54);
            Paragraph paragraph56 = body1.Elements<Paragraph>().ElementAt(55);
            Paragraph paragraph57 = body1.Elements<Paragraph>().ElementAt(56);
            Paragraph paragraph58 = body1.Elements<Paragraph>().ElementAt(57);
            Paragraph paragraph59 = body1.Elements<Paragraph>().ElementAt(58);
            SectionProperties sectionProperties1 = body1.GetFirstChild<SectionProperties>();
            paragraph1.RsidParagraphAddition = "0018454D";
            paragraph1.RsidRunAdditionDefault = "00545236";

            Run run1 = paragraph1.GetFirstChild<Run>();

            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            paragraph1.InsertBefore(bookmarkStart1, run1);

            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };
            paragraph1.InsertBefore(bookmarkEnd1, run1);
            run1.RsidRunProperties = null;

            Paragraph paragraph60 = new Paragraph() { RsidParagraphAddition = "00BA2F0F", RsidParagraphProperties = "00BA2F0F", RsidRunAdditionDefault = "00545236" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            SectionProperties sectionProperties2 = new SectionProperties() { RsidR = "00BA2F0F" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties2.Append(pageSize1);
            sectionProperties2.Append(pageMargin1);
            sectionProperties2.Append(columns1);
            sectionProperties2.Append(docGrid1);

            paragraphProperties1.Append(sectionProperties2);

            paragraph60.Append(paragraphProperties1);
            body1.InsertBefore(paragraph60, paragraph2);

            Paragraph paragraph61 = new Paragraph() { RsidParagraphAddition = "00BA2F0F", RsidParagraphProperties = "00BA2F0F", RsidRunAdditionDefault = "00545236" };

            Run run2 = new Run();
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text1 = new Text();
            text1.Text = "This is the paragraph two";

            run2.Append(lastRenderedPageBreak1);
            run2.Append(text1);

            paragraph61.Append(run2);
            body1.InsertBefore(paragraph61, paragraph2);

            Paragraph paragraph62 = new Paragraph() { RsidParagraphAddition = "00BA2F0F", RsidParagraphProperties = "00BA2F0F", RsidRunAdditionDefault = "00545236" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();

            SectionProperties sectionProperties3 = new SectionProperties() { RsidR = "00BA2F0F", RsidSect = "00BA2F0F" };
            PageSize pageSize2 = new PageSize() { Width = (UInt32Value)15840U, Height = (UInt32Value)12240U, Orient = PageOrientationValues.Landscape };
            PageMargin pageMargin2 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns2 = new Columns() { Space = "720" };
            DocGrid docGrid2 = new DocGrid() { LinePitch = 360 };

            sectionProperties3.Append(pageSize2);
            sectionProperties3.Append(pageMargin2);
            sectionProperties3.Append(columns2);
            sectionProperties3.Append(docGrid2);

            paragraphProperties2.Append(sectionProperties3);

            paragraph62.Append(paragraphProperties2);
            body1.InsertBefore(paragraph62, paragraph2);
            paragraph2.RsidParagraphAddition = "0018454D";
            paragraph2.RsidRunAdditionDefault = "00545236";

            ParagraphProperties paragraphProperties3 = paragraph2.GetFirstChild<ParagraphProperties>();

            Run run3 = new Run();
            LastRenderedPageBreak lastRenderedPageBreak2 = new LastRenderedPageBreak();
            Text text2 = new Text();
            text2.Text = "This is Paragraph three.";

            run3.Append(lastRenderedPageBreak2);
            run3.Append(text2);
            paragraph2.InsertBefore(run3, paragraphProperties3);

            paragraphProperties3.Remove();

            SectionProperties sectionProperties4 = new SectionProperties() { RsidR = "0018454D" };
            PageSize pageSize3 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin3 = new PageMargin() { Top = 1440, Right = (UInt32Value)1800U, Bottom = 1440, Left = (UInt32Value)1800U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns3 = new Columns() { Space = "720" };
            DocGrid docGrid3 = new DocGrid() { LinePitch = 360 };

            sectionProperties4.Append(pageSize3);
            sectionProperties4.Append(pageMargin3);
            sectionProperties4.Append(columns3);
            sectionProperties4.Append(docGrid3);
            body1.InsertBefore(sectionProperties4, paragraph3);

            paragraph3.Remove();
            paragraph4.Remove();
            paragraph5.Remove();
            paragraph6.Remove();
            paragraph7.Remove();
            paragraph8.Remove();
            paragraph9.Remove();
            paragraph10.Remove();
            paragraph11.Remove();
            paragraph12.Remove();
            paragraph13.Remove();
            paragraph14.Remove();
            paragraph15.Remove();
            paragraph16.Remove();
            paragraph17.Remove();
            paragraph18.Remove();
            paragraph19.Remove();
            paragraph20.Remove();
            paragraph21.Remove();
            paragraph22.Remove();
            paragraph23.Remove();
            paragraph24.Remove();
            paragraph25.Remove();
            paragraph26.Remove();
            paragraph27.Remove();
            paragraph28.Remove();
            paragraph29.Remove();
            paragraph30.Remove();
            paragraph31.Remove();
            paragraph32.Remove();
            paragraph33.Remove();
            paragraph34.Remove();
            paragraph35.Remove();
            paragraph36.Remove();
            paragraph37.Remove();
            paragraph38.Remove();
            paragraph39.Remove();
            paragraph40.Remove();
            paragraph41.Remove();
            paragraph42.Remove();
            paragraph43.Remove();
            paragraph44.Remove();
            paragraph45.Remove();
            paragraph46.Remove();
            paragraph47.Remove();
            paragraph48.Remove();
            paragraph49.Remove();
            paragraph50.Remove();
            paragraph51.Remove();
            paragraph52.Remove();
            paragraph53.Remove();
            paragraph54.Remove();
            paragraph55.Remove();
            paragraph56.Remove();
            paragraph57.Remove();
            paragraph58.Remove();
            paragraph59.Remove();
            sectionProperties1.Remove();
        }

        public void CreateOrientation(string filePath)
        {
            using (package = WordprocessingDocument.Open(filePath, true))
            {
                //package.AddMainDocumentPart();
                //package.MainDocumentPart.Document = new Document();
                Body body = package.MainDocumentPart.Document.Body; //package.MainDocumentPart.Document.Body = new Body();

                Paragraph para1 = CreatePara1(); // Paragraph 1 Portrait

                Paragraph breakBefore = CreateBreakBefore(); // Break before next paragraph that would be landscape

                Paragraph para2 = CreatePara2(); // Paragraph 2 Landscape

                Paragraph breakAfter = CreateBreakAfter(); // Break after previous paragraph that was landscape

                Paragraph para3 = CreatePara3(); // Paragraph 3 Portrait again

                body.Append(para1);
                body.Append(breakBefore);
                body.Append(para2);
                body.Append(breakAfter);
                body.Append(para3);

            }
        }
        private static Paragraph CreatePara1()
        {
            Paragraph para1 = new Paragraph();
            Run run1 = new Run(new Text("This is Paragraph one."));
            para1.Append(run1);
            return para1;
        }

        private static Paragraph CreateBreakBefore()
        {
            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00BA2F0F", RsidParagraphProperties = "00BA2F0F", RsidRunAdditionDefault = "00BA2F0F" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "00BA2F0F" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            paragraphProperties1.Append(sectionProperties1);

            paragraph2.Append(paragraphProperties1);
            return paragraph2;
        }

        private static Paragraph CreatePara2()
        {
            Paragraph para2 = new Paragraph() { RsidParagraphAddition = "00BA2F0F", RsidParagraphProperties = "00BA2F0F", RsidRunAdditionDefault = "00BA2F0F" };
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };

            Run run2 = new Run();
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text2 = new Text();
            text2.Text = "This is the paragraph two";

            run2.Append(lastRenderedPageBreak1);
            run2.Append(text2);

            para2.Append(bookmarkStart1);
            para2.Append(run2);
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };
            return para2;
        }

        private static Paragraph CreateBreakAfter()
        {
            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00BA2F0F", RsidParagraphProperties = "00BA2F0F", RsidRunAdditionDefault = "00BA2F0F" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();

            SectionProperties sectionProperties2 = new SectionProperties() { RsidR = "00BA2F0F", RsidSect = "00BA2F0F" };
            PageSize pageSize2 = new PageSize() { Width = (UInt32Value)15840U, Height = (UInt32Value)12240U, Orient = PageOrientationValues.Landscape };
            PageMargin pageMargin2 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns2 = new Columns() { Space = "720" };
            DocGrid docGrid2 = new DocGrid() { LinePitch = 360 };

            sectionProperties2.Append(pageSize2);
            sectionProperties2.Append(pageMargin2);
            sectionProperties2.Append(columns2);
            sectionProperties2.Append(docGrid2);

            paragraphProperties2.Append(sectionProperties2);

            paragraph4.Append(paragraphProperties2);
            return paragraph4;
        }

        private static Paragraph CreatePara3()
        {
            Paragraph para3 = new Paragraph();
            Run run3 = new Run(new Text("This is Paragraph three."));
            para3.Append(run3);
            return para3;
        }
    }
}
