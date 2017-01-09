using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using log4net.Repository.Hierarchy;
using log4net.Appender;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
namespace WordAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ShapeFormat_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void ParaIndentbtn_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (Footnote rngWord in Globals.ThisAddIn.Application.ActiveDocument.Content.Footnotes)
                rngWord.Range.ParagraphFormat.TabHangingIndent(Int16.Parse((rngWord.Range.ParagraphFormat.FirstLineIndent == 0 ? 1 : -1).ToString()));
                
        }

        private void addContentControl_Click(object sender, RibbonControlEventArgs e)
        {
            Selection selection = Globals.ThisAddIn.Application.Selection;
            //content control
            selection.Range.ContentControls.Add(WdContentControlType.wdContentControlRichText);
            selection.MoveRight();
            selection.TypeParagraph();
            //Legacy Forms controls
            selection.FormFields.Add(selection.Range,WdFieldType.wdFieldFormDropDown);
            selection.TypeParagraph();
            //ActiveX Controls
            selection.InlineShapes.AddOLEControl("Forms.ComboBox.1");
        }

        private void WrapTable_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Rows.WrapAroundText = -1;
        }

        private void SelectionMove_Click(object sender, RibbonControlEventArgs e)
        {
            Selection s = Globals.ThisAddIn.Application.Selection;
            Table t = Globals.ThisAddIn.Application.ActiveDocument.Tables[1];
            t.Select();
            s.MoveUp();
            MessageBox.Show(s.Text);
            s.MoveLeft();
            s.MoveDown(Unit:WdUnits.wdLine,Count:1,Extend:WdMovementType.wdExtend);
            MessageBox.Show(s.Text);
            s.MoveRight();
            s.MoveDown(Unit: WdUnits.wdLine, Count: 1, Extend: WdMovementType.wdExtend);
            MessageBox.Show(s.Text);
        }
        readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private void LogBtn_Click(object sender, RibbonControlEventArgs e)
        {
            string path = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
            MessageBox.Show(path);
            logger.Info("Hello World");
        }

        private void ChangeLogPath_Click(object sender, RibbonControlEventArgs e)
        {
            // Get the Hierarchy object that organizes the loggers
            Hierarchy hier = log4net.LogManager.GetRepository() as Hierarchy;

            if (hier != null)
            {
                // Get ADONetAppender
                var rollingFileAppender =
                    (RollingFileAppender)hier.GetAppenders().Where(
                        appender => appender.Name.Equals("RollingLogFileAppender", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();

                if (rollingFileAppender != null)
                {
                    rollingFileAppender.File = @"D:\log1\log.txt";
                    rollingFileAppender.ActivateOptions();
                }
            }
            MessageBox.Show("ok");
            logger.Info("Hello World");
        }

        private void changeLocation_Click(object sender, RibbonControlEventArgs e)
        {
            int x = Globals.ThisAddIn.Application.Left;
            int y = Globals.ThisAddIn.Application.Top;
            MessageBox.Show("Left is " +x.ToString()+" Top is "+y.ToString());
        }

        private void findReplace_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument; //this.Application.ActiveDocument;
            Word.Range rnge = document.Content;
            string findtxt = "Doc";
            Int16 intFound = 0;
            string OriginalText = "";
            string DocNum = "";
            //OfficeSVC.OfficeSvcClient Ofce = new OfficeSVC.OfficeSvcClient();
            string hpl = "";
            Int32 r;

            rnge.Find.ClearFormatting();
            rnge.Find.Text = "Doc";
            rnge.Find.Forward = true;
            var missing = Type.Missing;
            rnge.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing,
              ref missing, ref missing, ref missing, ref missing, ref missing);
            int i = 0;
            while (rnge.Find.Found)
            {
                rnge.Font.ColorIndex = WdColorIndex.wdBlue;
                i++;
                Word.Range rnge2 = document.Range(rnge.Start, rnge.End);
                rnge2.MoveEndWhile("0", ref missing);
                DocNum = rnge2.Text;
                OriginalText = rnge2.Text;
                DocNum = DocNum.Replace("Doc", "");
                if (int.TryParse(DocNum, out r) == true)
                    hpl = "http://" + DocNum.ToString();  // makes the hyperlink 
                if (hpl != "")
                {
                    object oAddress = hpl;
                    rnge2.Hyperlinks.Add(rnge2, oAddress, ref missing, ref missing, ref missing, ref missing);

                }

                //rnge2.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                //rnge2.Text = rnge2.Text.Replace("Doc",  i.ToString());
                //rnge2.MoveEndWhile(" 0123456789", ref missing); // extend the range to get the document #
                //DocNum = rnge2.Text;
                //OriginalText = rnge2.Text;
                //DocNum = DocNum.Replace("Doc", "Doc" + i.ToString());
                //if (int.TryParse(DocNum, out r) == true)
                //    hpl = Ofce.GetDocURL("3", "11", "cv", "146", DocNum);  // makes the hyperlink 
                //if (hpl != "")
                //{
                //    object oAddress = hpl;
                //    rnge2.Hyperlinks.Add(rnge2, oAddress, ref missing, ref missing, ref missing, ref missing);

                //}
        
                //rnge = document.Range(rnge.End,document.Application.Selection.GoTo(missing,WdGoToDirection.wdGoToLast,missing,missing).End);
                rnge.Find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

                //rnge.Find.Execute(
                //ref missing, ref missing, ref missing, ref missing, ref missing,
                //ref missing, ref missing, ref missing, ref missing, ref missing,
                //ref missing, ref missing, ref missing, ref missing, ref missing);
            }
        }

        private void InsertContentControl_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc=Globals.ThisAddIn.Application.ActiveDocument;
            Range myRange = Globals.ThisAddIn.Application.Selection.Range;
            ContentControl cc= doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, myRange);
            cc.Appearance = WdContentControlAppearance.wdContentControlTags;
            Globals.ThisAddIn.Application.Selection.MoveRight(WdUnits.wdCharacter,1,Type.Missing);
            Globals.ThisAddIn.Application.Selection.InsertAfter("");            
        }

        private void RangeReplace_Click(object sender, RibbonControlEventArgs e)
        {
            Range r = Globals.ThisAddIn.Application.Selection.Range;
            r.Text = "";
        }

        private void TableCell_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string heading = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading1].NameLocal;
            var missing=Type.Missing;          
            foreach (Word.Table table in doc.Tables)
            {
                foreach (Word.Paragraph paragraph in table.Range.Paragraphs)
                {
                    paragraph.Range.Select();
                    Globals.ThisAddIn.Application.Selection.Find.ClearFormatting();
                    Globals.ThisAddIn.Application.Selection.Find.set_Style("Heading 1 Char");
                    bool b = Globals.ThisAddIn.Application.Selection.Find.Execute(ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing
                            , true, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing);        
                    if (b==true)
                    {
                        paragraph.Range.Font.ColorIndex = WdColorIndex.wdBlue; //MessageBox.Show(paragraph.); //Debug.WriteLine(paragraph);
                    }
                }
            }
        }

        private void AddInName_Click(object sender, RibbonControlEventArgs e)
        {
            string name = System.IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().FullName);
            MessageBox.Show(name);
        }

        private void SaveAsTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            object FileName = @"C:\Users\v-tazho\Desktop\myfile.dotx";
            object FileFormat = Word.WdSaveFormat.wdFormatXMLTemplate;
            object LockComments = false;
            object AddToRecentFiles = true;
            object ReadOnlyRecommended = false;
            object EmbedTrueTypeFonts = false;
            object SaveNativePictureFormat = true;
            object SaveFormsData = true;
            object SaveAsAOCELetter = false;
            object Encoding = MsoEncoding.msoEncodingUSASCII;
            object InsertLineBreaks = false;
            object AllowSubstitutions = false;
            object LineEnding = Word.WdLineEndingType.wdCRLF;
            object AddBiDiMarks = false;
            object wdCompatibilityMode = 15;
            object missing = Type.Missing;
            //var a = Globals.ThisDocument.Application.ActiveDocument.Content.Text;
            //Globals.ThisDocument.Application.ActiveDocument.SaveAs2(ref FileName, ref FileFormat, ref LockComments,
            //ref missing, ref AddToRecentFiles, ref missing,
            //ref ReadOnlyRecommended, ref EmbedTrueTypeFonts,
            //ref SaveNativePictureFormat, ref SaveFormsData,
            //ref SaveAsAOCELetter, ref missing, ref missing,
            //ref missing, ref missing, ref missing, ref wdCompatibilityMode);
            Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(ref FileName, ref FileFormat, ref LockComments,
            ref missing, ref AddToRecentFiles, ref missing,
            ref ReadOnlyRecommended, ref EmbedTrueTypeFonts,
            ref SaveNativePictureFormat, ref SaveFormsData,
            ref SaveAsAOCELetter, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref wdCompatibilityMode);
          
            MessageBox.Show("ok");
        }
        private void headerFooter_Click(object sender, RibbonControlEventArgs e)
        {
            for (int i = 0; i < 10000; i++)
            {
                Word.Sections wdSectionCollection = Globals.ThisAddIn.Application.ActiveDocument.Sections;
                Word.Section wdFirstSection = wdSectionCollection[1];
                Word.HeadersFooters wdHeaderFooterCollection = wdFirstSection.Headers;
                Word.HeaderFooter wdHeaderFooter = wdHeaderFooterCollection[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                Marshal.FinalReleaseComObject(wdHeaderFooter);
                Marshal.FinalReleaseComObject(wdHeaderFooterCollection);
                Marshal.FinalReleaseComObject(wdFirstSection);
                Marshal.FinalReleaseComObject(wdSectionCollection);
            }
            GC.Collect();
        }

        private void MoveShape_Click(object sender, RibbonControlEventArgs e)
        {

            Word.Shape item = Globals.ThisAddIn.Application.ActiveDocument.Shapes[1];
            item.Select();
            Globals.ThisAddIn.Application.Selection.Copy();
            Globals.ThisAddIn.Application.Selection.Paste();
            Word.Shape newShape = Globals.ThisAddIn.Application.ActiveDocument.Shapes[2];
            InlineShape newshape = newShape.ConvertToInlineShape();
            string newxml = newshape.Range.WordOpenXML;
            InlineShape shape = item.ConvertToInlineShape();
            string xml = shape.Range.WordOpenXML;
            newShape.Delete();
            //float left = item.Left;
            //float top = item.Top;
            //InlineShape shape = item.ConvertToInlineShape();
            //string xml = shape.Range.WordOpenXML;
            //Word.Shape item1 = shape.ConvertToShape();
            //shape.Select();
            //Globals.ThisAddIn.Application.Selection.Application.Left =Convert.ToInt16( left);
            //Globals.ThisAddIn.Application.Selection.Application.Top = Convert.ToInt16(top);
            //item1.Left = left;
            //item1.Top = top;
        }

        private void InsertXMLbtn_Click(object sender, RibbonControlEventArgs e)
        {
            string xmlString =
    "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
    "<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'>" +
            "<w:r>" +
            "<w:rPr>" +
            "<w:color w:val='Black' /><w:rFonts w:ascii='Calibri' w:h-ansi='Calibri' w:cs='Calibri'/>" +
            "<w:sz w:val='22'/>" +
        "</w:rPr>" +
    "<w:fldChar w:fldCharType='begin' />"+
    "<w:instrText xml:space='preserve'> MacroButton TEST TEST</w:instrText>"+
                "<w:fldChar w:fldCharType='begin'>"+
                "<w:fldData xml:space='preserve'>KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq"+
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA"+
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq"+
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA"+
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq"+
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA"+
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq"+
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA"+
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq"+
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA"+
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq"+
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKg =</w:fldData>" +
                "</w:fldChar>" +
                " <w:instrText xml:space='preserve'> ADDIN TEST</w:instrText>" +
                " <w:fldChar w:fldCharType='end' />" +
                "<w:fldChar w:fldCharType='end' />" +
                "</w:r>" +
                "</w:p>";
            xmlString =  System.IO.File.ReadAllText(@"D:\Edward\Project\MSDNProject\MSDNProject\WordAddIn\Field.xml");
            Globals.ThisAddIn.Application.Selection.InsertXML(xmlString);

        }
        Field f;
        private void InsertAddInField_Click(object sender, RibbonControlEventArgs e)
        {
            Document d = Globals.ThisAddIn.Application.ActiveDocument;
            Range r = Globals.ThisAddIn.Application.Selection.Range;
            f = Globals.ThisAddIn.Application.ActiveDocument.Fields.Add(r,WdFieldType.wdFieldAddin,"Text");
            string data="KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq" +
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA" +
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq" +
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA" +
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq" +
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA" +
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq" +
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA" +
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq" +
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoA" +
"KgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAq" +
"ACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKgAqACoAKg";
            MessageBox.Show(data.Length.ToString());
            f.Data = data;
        }

        private void getAddInField_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show(f.Data);
            foreach (Field ff in Globals.ThisAddIn.Application.ActiveDocument.Fields)
            {
                string s = Globals.ThisAddIn.Application.ActiveDocument.Fields[1].Data;
            
                MessageBox.Show(s);
            }
        }

        private void WordArtbtn_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (Microsoft.Office.Interop.Word.Shape shape in Globals.ThisAddIn.Application.ActiveDocument.InlineShapes)
            {
                System.Diagnostics.Debug.WriteLine(shape.AlternativeText);
            }
            //float leftPosition = (float)Globals.ThisAddIn.Application.Selection.Information[
            //    Microsoft.Office.Interop.Word.WdInformation.wdHorizontalPositionRelativeToPage];
            //float topPosition = (float)Globals.ThisAddIn.Application.Selection.Information[
            //    Microsoft.Office.Interop.Word.WdInformation.wdVerticalPositionRelativeToPage];
            //Microsoft.Office.Interop.Word.Shape shape = Globals.ThisAddIn.Application.ActiveDocument.Shapes.AddTextEffect(
            //   Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect29, "SampleText",
            //   "Arial Black", 24, Microsoft.Office.Core.MsoTriState.msoFalse,
            //   Microsoft.Office.Core.MsoTriState.msoFalse, leftPosition, topPosition);
            ////shape.Select();
            ////var rng =  // Globals.ThisAddIn.Application.Selection.ShapeRange;
            ////shape.Left = (float)WdShapePosition.wdShapeCenter;
            ////shape.Top = (float)WdShapePosition.wdShapeCenter;
            ////shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            //shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            //shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
            
            //shape.Left = (int)WdShapePosition.wdShapeCenter;
            //shape.Top = (int)WdShapePosition.wdShapeCenter;
            //shape.TextFrame2.WarpFormat = MsoWarpFormat.msoWarpFormat15;
            //shape.LayoutInCell = 0;
            //rng.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            //rng.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            //rng.RelativeHorizontalSize = WdRelativeHorizontalSize.wdRelativeHorizontalSizePage;
            //rng.RelativeVerticalSize = WdRelativeVerticalSize.wdRelativeVerticalSizePage;
            //rng.Left = (float)WdShapePosition.wdShapeCenter;
            //  Globals.ThisAddIn.Application.Selection.ShapeRange.LeftRelative = (float)wdShapePositionRelativeNone;
            //rng.Top = (float)WdShapePosition.wdShapeCenter;
            //rng.TopRelative = (float)WdShapePositionRelative. wdShapePositionRelativeNone;
            //rng.WidthRelative = (float)WdShapeSizeRelative. wdShapeSizeRelativeNone;
            //rng.HeightRelative = (float)WdShapeSizeRelative.wdShapeSizeRelativeNone;
            //rng.LockAnchor = 0;
            //rng.LayoutInCell =0;
        }

        private void Restrictionbtn_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            doc.Unprotect("123");
            doc.Styles["Line Number"].Locked = !doc.Styles["Line Number"].Locked;//true;  //disable heading 1
            //doc.Styles["heading 2"].Locked = false; //enable heading 2
            doc.Protect(WdProtectionType.wdNoProtection, Type.Missing, "123", Type.Missing, true);
        }

        private void paraID_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application oWord = Globals.ThisAddIn.Application;
            Word.Shape shNew = oWord.ActiveDocument.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
 100, 100, 100, 100);               
            shNew.Select();
            int id = 10 ;
            shNew.Name="tET";
            Word.Paragraphs ps = oWord.Selection.Paragraphs;
            ps[1].ID = "IND1";
            ps[1].Range.Text = "Test";
            ps[1].Range.ID = "IND1";
            
            foreach (Word.Paragraph p in oWord.ActiveDocument.Paragraphs)
            {
                p.ID = "NO" + id.ToString();
                id++;
            }
            //int id=1;
            //Word.Paragraphs ps = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs;
            //ps[1].ID = "IND1";
            //foreach(Word.Paragraph p in Globals.ThisAddIn.Application.ActiveDocument.Paragraphs)
            //{
            //    p.ID = "NO" + id.ToString();
            //}
        }

        private void SetFont_Click(object sender, RibbonControlEventArgs e)
        {
            Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Paragraph pp = document.Paragraphs.Add(Type.Missing);
            pp.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            pp.Range.Font.Name = "B Nazanin";
            pp.Range.Font.Size = 16;
            pp.Range.InsertParagraphAfter();
            pp.Range.Text = "Arabic Text سلام "; 

        } 
    }
}
