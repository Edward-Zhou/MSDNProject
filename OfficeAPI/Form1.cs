using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAPI.Class;
using OfficeAPI.Class.OpenXmlHelper;
using OfficeAPI.Class.ExcelAutomation;
using OfficeAPI.Class.WordAutomation;
using OfficeAPI.Class.OpenXmlHelper.WordHelper;
using OfficeAPI.Class.OpenXmlHelper.PPTHelper;
using excel=Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Runtime.InteropServices;
using OfficeAPI.Class.OpenXmlHelper.ExcelHelper;
using outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Excel;
using OneNote = Microsoft.Office.Interop.OneNote;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Excel.Utilities;
using OfficeAPI.Class.OpenXmlHelper.PPTHelper;
using System.Net.Mail;
using System.Net;
using _word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace OfficeAPI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string fileName;
        private void StyleSet_Click(object sender, EventArgs e)
        {
            fileName = @"D:\OfficeDev\OpenXML\Font.docx";
            //var node = ExtractStylesPart(fileName);            
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                var docPart = document.MainDocumentPart;
                StyleDefinitionsPart stylesPart = null;
                stylesPart = docPart.StyleDefinitionsPart;
                ChangeStyleDefinitionsPart1(stylesPart);                
            }
        }
        public static XDocument ExtractStylesPart(string fileName)
        {
            XDocument styles = null;
            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                var docPart = document.MainDocumentPart;
                StylesPart stylesPart = null;
                stylesPart = docPart.StyleDefinitionsPart;
                if (stylesPart != null)
                {
                    using (var reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        styles = XDocument.Load(reader);
                    }
                }
            }
            return styles;
        }
        private static void ChangeStyleDefinitionsPart1(StyleDefinitionsPart styleDefinitionsPart1)
        {
            //DocumentFormat.OpenXml.Wordprocessing.Styles styles1 = styleDefinitionsPart1.Styles;

            //DocumentFormat.OpenXml.Wordprocessing.Style style1 = styles1.GetFirstChild<Style>(); //get the specifc style

            //Rsid rsid1 = new Rsid() { Val = "00B10D4B" };
            //style1.Append(rsid1);

            //StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            //FontSize fontSize1 = new FontSize() { Val = "144" };

            //styleRunProperties1.Append(fontSize1);
            //style1.Append(styleRunProperties1);
        }

        private void ExcelHeaderImg_Click(object sender, EventArgs e)
        {
            ExcelImage.CreateSummaryExcelDoc(@"C:\Users\Administrator\Desktop\Testnew.xlsx");
            MessageBox.Show("ok");
        }

        private void Savebtn_Click(object sender, EventArgs e)
        {
            string xlsxFn = @"D:\Edward\Project\MSDNProject\MSDNProject\OfficeAPI\Test.xlsx";
            string xlsFn = @"D:\Edward\Project\MSDNProject\MSDNProject\OfficeAPI\Test1.xls";
            var Excel = new Microsoft.Office.Interop.Excel.Application();
            var wkb = Excel.Workbooks.Open(xlsxFn, false, true);
            wkb.SaveAs(xlsFn, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8);
            wkb.Close();
            Excel.Quit();
        }

        private void CreateShape_Click(object sender, EventArgs e)
        {
            string xlsxFn = @"D:\Edward\Project\MSDNProject\MSDNProject\OfficeAPI\Test.xlsx";
            var Excel = new Microsoft.Office.Interop.Excel.Application();
            Excel.Visible = true;
            var wkb = Excel.Workbooks.Open(xlsxFn, false, true);

            Microsoft.Office.Interop.Excel.Shape btn2 = wkb.ActiveSheet.Shapes.AddFormControl(Microsoft.Office.Interop.Excel.XlFormControl.xlButtonControl, 150, 5, 150, 22);
            btn2.ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset1;
            btn2.Name = "Update";
            //ExcelAutomation.LOCK_CELLS_PROTECT_CONTENTS(wkb.ActiveSheet);
        }

        private void ExcelSharedString_Click(object sender, EventArgs e)
        {
            ExcelHelp.InsertText(@"C:\Users\v-tazho\Desktop\Test.xlsx", "Inserted Text");
            MessageBox.Show("ok");
        }

        private void ExcelDeleteFormula_Click(object sender, EventArgs e)
        {
            DeleteFormula df = new DeleteFormula();
            df.ChangePackage(@"D:\OfficeDev\OpenXML\Formula - Copy.xlsx");
            MessageBox.Show("ok");
        }

        private void PPTSlide_Click(object sender, EventArgs e)
        {
            //InsertSlide Islide=new InsertSlide();
            //Islide.InsertNewSlide(@"D:\OfficeDev\PPT\LayoutTest.pptx", 1, "My new slide");
            InsertSlide.InsertNewSlide(@"D:\OfficeDev\PPT\LayoutTestOK.pptx", 1, "Comparison");
            //InsertSlideTest.InsertNewSlide(@"D:\OfficeDev\PPT\LayoutTest - Copy.pptx", 1, "Comparison");
            //InsertSlideMe ism = new InsertSlideMe();
            //ism.ChangePackage(@"D:\OfficeDev\PPT\LayoutTestOK.pptx"); 
            MessageBox.Show("ok");
        }

        private void GoToBookMark_Click(object sender, EventArgs e)
        {
            WordAutomation.goToBookMark();
        }

        private void GetTableCount_Click(object sender, EventArgs e)
        {
            GetTableCountClass.getTableCount();
        }

        private void setOptions_Click(object sender, EventArgs e)
        {
            setOptionClass.setOptions();
        }

        private void ChangeSheetName_Click(object sender, EventArgs e)
        {
            SheetName.TriggerExcel2010Bug();
        }

        private void ExcelAddFormula_Click(object sender, EventArgs e)
        {
            AddFormula addFormula = new AddFormula();
            addFormula.ChangePackage(@"D:\OfficeDev\OpenXML\FormulaNoBak.xlsx");
            MessageBox.Show("ok");
        }

        private void DateFormat_Click(object sender, EventArgs e)
        {
            DateFormatExcel dfe = new DateFormatExcel();
            dfe.ChangePackage(@"D:\OfficeDev\OpenXML\StyleIndex.xlsx");
            MessageBox.Show("ok");
        }

        private void getCell_Click(object sender, EventArgs e)
        {
            GetCellValue getcell = new GetCellValue();
            getcell.getCell(@"D:\OfficeDev\OpenXML\Excel\20160822\ExcelRangeTarget.xlsx");
        }

        private void DateRegionFormat_Click(object sender, EventArgs e)
        {
            DateRegion dr = new DateRegion();
            dr.ChangePackage(@"DateFormatChangeBak - Copy - Copy.xlsx");
            MessageBox.Show("ok");
        }

        private void setCell_Click(object sender, EventArgs e)
        {
            setCellValue scv = new setCellValue();
            //scv.setCell(@"D:\OfficeDev\OpenXML\DateFormatChange - Copy (3).xlsx");
            scv.setCell(@"D:\OfficeDev\OpenXML\Excel\Empty - Copy (2).xlsx");
            MessageBox.Show("ok");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
  
        }


        private void ParagraphMovebtn_Click(object sender, EventArgs e)
        {
            ParagraphMove pm = new ParagraphMove();
            pm.paragraphMove(@"D:\OfficeDev\OpenXML\Word\ParaCopy1.docx");
            MessageBox.Show("ok");
        }

        private void CreateDocbtn_Click(object sender, EventArgs e)
        {
            CreateDoc dc = new CreateDoc();
            dc.CreateDocFromTem();
            MessageBox.Show("ok");
        }

        private void AddNotebtn_Click(object sender, EventArgs e)
        {
            string filePath = @"D:\OfficeDev\OpenXML\PPT\Empty - Copy.pptx";

            AddNoteClass.AddNote(filePath, 0);
           // AddNoteClass.ChangePresentationPart();
            MessageBox.Show("ok");
        }
        Microsoft.Office.Interop.Outlook.Application oApp;
        Microsoft.Office.Interop.Outlook.Explorer oEx;
        Microsoft.Office.Interop.Outlook.Selection oSel;
        Microsoft.Office.Interop.Outlook.MailItem oMail;
        Microsoft.Office.Interop.Outlook.NameSpace oName;
        private void button1_Click_1(object sender, EventArgs e)
        {
           oApp = new Microsoft.Office.Interop.Outlook.Application();
           oEx = oApp.ActiveExplorer();
           oEx.Display();
           oSel = oEx.Selection;
           oMail = oSel[1];
           oName=oApp.GetNamespace("MAPI");
           Microsoft.Office.Interop.Outlook.MAPIFolder folder = oName.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox).Folders["Tina"];
           oMail.Move(folder);
           MessageBox.Show(oMail.Subject);
            //Export(@"C:\Users\v-tazho\Desktop\Test.xlsx", "Sheet1", 1, @"C:\Users\v-tazho\Desktop\Test.gif", "gif");
        }

        public void Export(string fileName, string sheetName, int chartIndex, string targetFile, string filter)
        {
            //string strCmdText;
            //strCmdText = "excel /x ";
            //System.Diagnostics.Process.Start("CMD.exe", strCmdText);
            //Console.WriteLine(Thread.CurrentThread.CurrentCulture.Name);
            excel.Application excelApp = new excel.Application();
            excelApp.Visible = true;
            excel.Workbook aWorkbook = excelApp.Workbooks.Open(fileName);

            //Microsoft.Office.Interop.Excel.Chart aChart = aWorkbook.Worksheets[sheetName].ChartObjects[chartIndex].Chart;
            //Console.WriteLine("Convert successfully: {0}", aChart.Export(targetFile, filter));

        }

        private void FormProtectbtn_Click(object sender, EventArgs e)
        {
            FormProtect fp = new FormProtect();
            fp.formProtect(@"D:\OfficeDev\OpenXML\Word\WordProtectionNo - Copy.docx");
            MessageBox.Show("ok");
        }

        private void FormProtectTotalbtn_Click(object sender, EventArgs e)
        {
            FormProTotal fp = new FormProTotal();
            fp.ChangePackage(@"D:\OfficeDev\OpenXML\Word\WordProtectionNo - Copy.docx");
            MessageBox.Show("ok");
        }

        private void RtfImportbtn_Click(object sender, EventArgs e)
        {
            RtfImport.ImportRtf(@"D:\OfficeDev\OpenXML\Word\ImportRtf.docx");
            MessageBox.Show("ok");
        }

        private void ExcelMovebtn_Click(object sender, EventArgs e)
        {

        }

        private void RemoveContentControlbtn_Click(object sender, EventArgs e)
        {
            DeleteContentControl dcc = new DeleteContentControl();
            dcc.RemoveContentControl(@"D:\OfficeDev\OpenXML\Word\ContentControl - Copy.docx");
        }

        private void ChangeOrientbtn_Click(object sender, EventArgs e)
        {
            Orientations ori = new Orientations();
            ori.ChangeOrientation(@"D:\OfficeDev\OpenXML\Word\WordOrientationsT3.docx");
            MessageBox.Show("ok");
        }

        private void ExternalPart_Click(object sender, EventArgs e)
        {
            GetRelationShip grt = new GetRelationShip();
            grt.getExternalRelationShip(@"D:\OfficeDev\OpenXML\Word\ExternalRelationship.docx");
        }

        private void OpenXlabtn_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //app.Workbooks.Add();
            //var add = app.AddIns.Add(@"D:\OfficeDev\Excel\XLA.xla");
            //add.Installed = true;
            app.Workbooks.Open(@"D:\OfficeDev\Excel\Testxla.xla");

            //app.Visible = true;
            try
            {
                //app.Application.Visible = true;
                //app.Visible = true;
                //app.UserControl = false;
                //app.Application.Visible = true;
                app.Visible = true;
                MessageBox.Show(app.Visible.ToString()); //false
            }
            catch { }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;

        }

        private void CellFormatbtn_Click(object sender, EventArgs e)
        {
            CellFormat1 cf = new CellFormat1();
            cf.cellFormat();
            MessageBox.Show("ok");
        }

        private void button3_Click(object sender, EventArgs e)
        {


        }

        private void Validation_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wk= app.Workbooks.Open(@"D:\OfficeDev\Excel\Test.xlsx");
            app.Visible = true;
            Worksheet ws = wk.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range r = ws.Range["C4"];
            r.Validation.Add( XlDVType.xlValidateDecimal, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, "0", "100");
            MessageBox.Show("ok");
        }

        private void AddBtn_Click(object sender, EventArgs e)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = @"C:\Users\v-tazho\Documents\Custom Office Templates\Getting Started Tutorial.dotx";

            //Create new application
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;
            word.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            //Activate document
            Microsoft.Office.Interop.Word.Document doc = word.ActiveDocument;

            doc.Activate();
        }

        private void getCellValue_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wk = app.Workbooks.Open(@"C:\Users\v-tazho\Desktop\Test.xlsx"); //excel file path
            app.Visible = true;
            Worksheet ws = wk.ActiveSheet;
            excel.Range r = ws.Range["B12"];
            MessageBox.Show(r.Text);
            string rSum = app.WorksheetFunction.Sum(ws.Range["A1:A5"]).ToString();
            MessageBox.Show(rSum);
        }

        private void HtmlToWord_Click(object sender, EventArgs e)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(@"D:\OfficeDev\OpenXML\Word\201512\Html.docx", true))
            {
                string altChunkId = "myId";
                MainDocumentPart mainDocPart = doc.MainDocumentPart;

                var run = new Run(new Text("test"));
                var p = new  DocumentFormat.OpenXml.Wordprocessing.Paragraph(new ParagraphProperties(
                     new Justification() { Val = JustificationValues.Center }),
                                   run);

                var body = mainDocPart.Document.Body;
                body.Append(p);

                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(@"<html><head></head><body><h1 style=""color:red"">HELLO</h1></body></html>"));      

                // Uncomment the following line to create an invalid word document.
                // MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes("<h1>HELLO</h1>"));

                // Create alternative format import part.
                AlternativeFormatImportPart formatImportPart =
                   mainDocPart.AddAlternativeFormatImportPart(
                      AlternativeFormatImportPartType.Html, altChunkId);
                //ms.Seek(0, SeekOrigin.Begin);

                // Feed HTML data into format import part (chunk).
                formatImportPart.FeedData(ms);
                AltChunk altChunk = new AltChunk();
                altChunk.Id = altChunkId;

                mainDocPart.Document.Body.Append(altChunk);
            }
        }

        private void HideAccessField_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Access.Application access = new Microsoft.Office.Interop.Access.Application();
            access.Visible=true;
            access.OpenCurrentDatabase(@"C:\Users\v-tazho\Documents\Test.accdb");
            Microsoft.Office.Interop.Access.Dao.Database dbs = access.CurrentDb();
            Microsoft.Office.Interop.Access.Dao.Field fld = dbs.TableDefs["cTime"].Fields["Test1"];
            fld.Properties["ColumnHidden"].Value = true;
        }
        ExcelLifeTimeManager excelManager;
        Worksheet activeWorkSheet;
        string findWhat = "DataToFind";
        string replaceWith = "DataToReplace";
        private void FindReplace_Click(object sender, EventArgs e)
        {
            activeWorkSheet.Range["A1:A2"].Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, false, false);
            //excelManager.Excel.Application.Selection.Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, false, false);
            //activeWorkSheet.Range["A1:A2"].Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true,false,false);
            
            //activeWorkSheet.Columns["a"].Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, false, false);
            //Range activeCell = excelManager.Excel.ActiveCell;
            //activeCell.Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true);
            //Range r = activeWorkSheet.Range[activeCell.Address];
            //r.Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true);
            //using (ExcelLifeTimeManager excelManager = new ExcelLifeTimeManager())
            //{
            //    Worksheet activeWorkSheet = excelManager.Excel.ActiveSheet as Worksheet;
            //    string findWhat = "DataToFind";
            //    string replaceWith = "DataToReplace";
            //    if (activeWorkSheet != null)
            //    {
            //        string[,] data = new string[3, 3];

            //        for (int outerIndex = 0; outerIndex < data.GetUpperBound(0); outerIndex++)
            //        {
            //            for (int innerIndex = 0; innerIndex < data.GetUpperBound(1); innerIndex++)
            //            {
            //                data[outerIndex, innerIndex] = findWhat;
            //            }
            //        }

            //        Range rangeToWriteData = activeWorkSheet.Range["A1", "C3"];
            //        rangeToWriteData.Value2 = data;
            //        Range activeCell = excelManager.Excel.ActiveCell;
            //        //Range r = activeWorkSheet.Range[activeCell.Address];
            //        //r.Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true);

            //        if (activeCell.Value == findWhat)
            //        {
            //            activeCell.Value = replaceWith;
            //        }
            //        //r.Select();
            //        //activeWorkSheet.Range["A5"].Value2 = activeCell.Address;
            //        //r.Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true);
            //        //activeWorkSheet.Range["A8"].Value2 = activeCell.Address;
            //        // Make sure active cell is A1
            //        //Assert.IsTrue((activeCell.Row == 1) && (activeCell.Column == 1) && 
            //        //    (excelManager.Excel.Cells.Count == 1));
            //        //Assert.IsTrue((activeCell.Row == 1) && (activeCell.Column == 1) );
            //        //activeWorkSheet.Range["A5"].Value2 = activeCell;
            //        //activeCell.Replace(findWhat, replaceWith, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true);
            //        //activeWorkSheet.Range["A10"].Value2 = activeCell;
            //        // We replaced only the active cell. We expect next occurence, so nextOccurence should not be null.
            //        //Range nextOccurence = activeWorkSheet.UsedRange.Find(findWhat, activeCell, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true);
            //        // Below assert is failing, since Range.Replace is replacing all the instances of search data in the worksheet.
            //        //Assert.IsNotNull(nextOccurence);
            //    }
            //}
        }
        /// <summary>
        /// Utility class to manage Excel instance.
        /// </summary>
        private class ExcelLifeTimeManager : IDisposable
        {
            internal Microsoft.Office.Interop.Excel.Application Excel { get; private set; }

            /// <summary>
            /// Creates instance of <see cref="ExcelLifeTimeManager"/> class.
            /// </summary>
            public ExcelLifeTimeManager()
            {
                Excel = new Microsoft.Office.Interop.Excel.Application();
                Excel.Visible = true;
                Excel.Workbooks.Add();
            }

            /// <summary>
            /// Clean up the resources.
            /// </summary>
            public void Dispose()
            {
                //foreach (Workbook workbook in Excel.Workbooks)
                //{
                //    workbook.Close(SaveChanges: false);
                //}
                //Excel.Quit();
                //Excel = null;
            }
        }

        private void CreateEmail_Click(object sender, EventArgs e)
        {
            outlook.Application oApp = new outlook.Application();
            outlook._MailItem oMailItem = (outlook._MailItem)oApp.CreateItem(outlook.OlItemType.olMailItem);
            oMailItem.To = "vv@hotmail.com";
            oMailItem.Subject = "Test";
            // body, bcc etc...
            oMailItem.Display(true);
           

        }

        private void FindReplaceExcel_Click(object sender, EventArgs e)
        {
            using (excelManager = new ExcelLifeTimeManager())
            {
                 activeWorkSheet = excelManager.Excel.ActiveSheet as Worksheet;

                if (activeWorkSheet != null)
                {
                    string[,] data = new string[3, 3];

                    for (int outerIndex = 0; outerIndex < data.GetUpperBound(0); outerIndex++)
                    {
                        for (int innerIndex = 0; innerIndex < data.GetUpperBound(1); innerIndex++)
                        {
                            data[outerIndex, innerIndex] = findWhat;
                        }
                    }
                  
                    excel.Range rangeToWriteData = activeWorkSheet.Range["A1", "C3"];
                    rangeToWriteData.Value2 = data;
                }
            }
        }

        private void AcceptRevisionsbtn_Click(object sender, EventArgs e)
        {
            string documentName = @"D:\OfficeDev\Word\201601\TrackedChanges_ON - Copy (2).docx";
            //AcceptRevisionsClass.AcceptRevisions(@"D:\OfficeDev\Word\201601\TrackedChanges_ON - Copy (2).docx", "ard21");
            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(documentName, false))
            {
                if (AcceptRevisionsClass.HasTrackedRevisions(wordDoc))
                    MessageBox.Show("{0} contains tracked revisions", documentName);
                else
                    MessageBox.Show("{0} does not contain tracked revisions", documentName);
            }

            MessageBox.Show("ok");
        }

        private void PivotTableFilter_Click(object sender, EventArgs e)
        {

        }

        private void DeleteCommand_Click(object sender, EventArgs e)
        {
            DeleteCommandbar dc = new DeleteCommandbar();
            dc.ChangePackage(@"D:\OfficeDev\Word\201602\Normalbak - Copy.dotm");
        }



        private void OneNotebtn_Click(object sender, EventArgs e)
        {
            String strXML;
            OneNote.Application onApplication = new OneNote.Application();

            var currentNotebookID = onApplication.Windows.CurrentWindow.CurrentNotebookId;
            onApplication.GetHierarchy(null,
                OneNote.HierarchyScope.hsPages, out strXML);

        }
        PowerPoint.Application pptApp;
        private void AddShapebtn_Click(object sender, EventArgs e)
        {
            string strTemplate = @"D:\OfficeDev\PPT\Paste.pptx";
            pptApp = new PowerPoint.Application();
            pptApp.Visible= Microsoft.Office.Core.MsoTriState.msoTrue;
            PowerPoint.Presentations pres = pptApp.Presentations;
            PowerPoint.Presentation pre = pres.Open(strTemplate,
        MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
            PowerPoint.Shapes shapes = pre.Slides[1].Shapes;
            shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,50, 50, 100, 200);

        }

        private void AddShapeToChart_Click(object sender, EventArgs e)
        {
            AddShapeToChartClass cp = new AddShapeToChartClass();
            cp.ChangePackage(@"D:\OfficeDev\OpenXML\Excel\UserShapes - Copy.xlsx");
        }

        private void AddShapeAuto_Click(object sender, EventArgs e)
        {
            PowerPoint.Application powerpoint = new PowerPoint.Application();
            var presentations = powerpoint.Presentations;
            PowerPoint.Presentation pres = presentations.Open(@"C:\Users\v-tazho\Documents\Presentation1.pptx", 
                                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                                        Microsoft.Office.Core.MsoTriState.msoFalse);

            try
            {
                //Instantiate slide object
                Microsoft.Office.Interop.PowerPoint.Slide objSlide = null;

                //Access the first slide of presentation
                objSlide = pres.Slides[1];

                PowerPoint.Chart ppChart = objSlide.Shapes.AddChart(Microsoft.Office.Core.XlChartType.xl3DColumn, 20F, 30F, 400F, 300F).Chart;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                //Close Workbook and presentation
                pres.Application.Quit();
            }
               

        }

        private void ExcelXML_Click(object sender, EventArgs e)
        {
           bool b= ExcelTemplateUtils.TemplateDownloadStuffWithData();
           MessageBox.Show("ok");
        }

        private void getCharacters_Click(object sender, EventArgs e)
        {
            RetrieveProperty.GetPropertyFromDocument(@"C:\Users\v-tazho\Downloads\OneDrive-2016-06-07\problemInOpenXML.docx");
            RetrieveProperty.getAllWords(@"C:\Users\v-tazho\Downloads\OneDrive-2016-06-07\problemInOpenXML.docx");
            //RetrieveProperty.GetPropertyFromDocument(@"C:\Users\v-tazho\Downloads\OneDrive-2016-06-07\problemInOpenXMLBak.docx");
        }

        private void PasteFromSourceFormatbtn_Click(object sender, EventArgs e)
        {
            pptApp.CommandBars.ExecuteMso("PasteSourceFormatting");
        }
      
        static int GetRunLength(XElement e)
        {
            return e
                .Descendants(W.t)
                .Select(t => (string)t)
                .StringConcatenate()
                .Length;
        }
        // return the run split locations for all runs in the paragraph
        static int[] RunSplitLocations(XElement paragraph)
        {
            // find the runs that don't have w:del or w:moveFrom as parent elements
            var runElements = paragraph
                .Descendants(W.r)
                .Where(e => e.Parent.Name != W.del && e.Parent.Name != W.moveFrom &&
                    e.Descendants(W.t).Any());

            // determine the run length of each run
            var runs = runElements
                .Select(r => new
                {
                    RunElement = r,
                    RunLength = GetRunLength(r)
                });

            // determine the split locations
            var runSplits = runs
                .Select(r => runs
                    .TakeWhile(a => a.RunElement != r.RunElement)
                    .Select(z => z.RunLength)
                    .Sum());

            return runSplits.ToArray();
        }

        // if value starts or ends with a space, return xml:space="preserve" attribute
        // else return null
        static XAttribute XmlSpacePreserved(string value)
        {
            if (value.Substring(0, 1) == " " || value.Substring(value.Length - 1) == " ")
                return new XAttribute(XNamespace.Xml + "space", "preserve");
            else
                return null;
        }

        private class RunSplits
        {
            public XElement RunElement { get; set; }
            public int RunLength { get; set; }
            public int RunLocation { get; set; }
        }

        private static object RunTransform(XElement element,
            int[] positions, IEnumerable<RunSplits> runSplits)
        {
            // split runs that have child text elements
            if (element.Name == W.r && element.Descendants(W.t).Any())
            {
                // get text of run
                string text = element
                    .Descendants(W.t)
                    .Select(t => (string)t).StringConcatenate();

                // find run in runSplits
                RunSplits rs = runSplits.First(r => r.RunElement == element);

                // find list of splits in this run
                var splitsInThisRun = positions
                    .Where(p => p >= rs.RunLocation && p < rs.RunLocation + rs.RunLength);

                // adjust splits so that split locations are relative to this run instead of
                // relative to the beginning of the paragraph
                var splitsIntext = splitsInThisRun
                    .Select(p => p - rs.RunLocation)
                    .ToArray();

                // project collection of strings that will be in the new, split runs
                var splitText = splitsIntext
                    .Select((p, i) =>
                        i != splitsIntext.Length - 1 ?
                        text.Substring(p, splitsIntext[i + 1] - p) :
                        text.Substring(p)
                );

                // project collection of runs that will replace the original run
                return splitText.Select(r =>
                    new XElement(W.r,
                        rs.RunElement.Elements().Where(e => e.Name != W.t),
                        new XElement(W.t,
                            XmlSpacePreserved(r),
                            r)));
            }

            // clone elements other than runs
            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n =>
                {
                    XElement e = n as XElement;
                    if (e != null)
                        return RunTransform(e, positions, runSplits);
                    return n;
                })
            );
        }

        public static XElement SplitRunsInParagraph(XElement p, int[] positions)
        {
            // find the runs that don't have w:del or w:moveFrom as parent elements
            var runElements = p
                .Descendants(W.r)
                .Where(e => e.Parent.Name != W.del && e.Parent.Name != W.moveFrom &&
                    e.Descendants(W.t).Any());

            // calculate the run length of each run
            var runs = runElements
                .Select(r => new
                {
                    RunElement = r,
                    RunLength = GetRunLength(r)
                });

            // calculate the location of each split
            var runSplits = runs
                .Select(r => new RunSplits
                {
                    RunElement = r.RunElement,
                    RunLength = r.RunLength,
                    RunLocation = runs
                        .TakeWhile(a => a.RunElement != r.RunElement)
                        .Select(z => z.RunLength)
                        .Sum()
                });

            // the positions argument contains a list of locations where splits will be added
            // to the paragraph.  In addition, runs may already be split at various places, and
            // we want those splits to remain, so we need to create the complete list of
            // locations where we want run splits.

            // create ordered union of desired splits and existing splits
            int[] allSplits = runSplits
                .Select(rs => rs.RunLocation)
                .Concat(positions)
                .OrderBy(s => s)
                .Distinct()
                .ToArray();

            // transform the paragraph to a new paragraph with new splits in runs
            return new XElement(W.p,
                p.Elements().Select(e => RunTransform(e, allSplits, runSplits))
            );
        }
        private void SplitRunsbtn_Click(object sender, EventArgs e)
        {
            using (WordprocessingDocument doc1 =
            WordprocessingDocument.Open(@"D:\OfficeDev\OpenXML\Word\201607\CommentCopy.docx", true))
            {
                XDocument doc = doc1.MainDocumentPart.GetXDocument();
                XElement p = doc.Root.Element(W.body).Element(W.p);
                //XElement newPara = SplitRunsInParagraph(p, new[] { 12, 15 });
                XElement newPara = SplitRunsInParagraph(p, new[] { 5,20 });
                p.ReplaceWith(newPara);
                doc1.MainDocumentPart.PutXDocument();
                Console.WriteLine(newPara);
            }

        }

        private void FileShape_Click(object sender, EventArgs e)
        {
            SetPPTShapeColorClass.SetPPTShapeColor(@"D:\OfficeDev\OpenXML\PPT\FillShape.pptx");
            MessageBox.Show("ok");
            
        }

        private void OpenBtn_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                
                string fileName = openFileDialog1.FileName;
               // using(SpreadsheetDocument ss=SpreadsheetDocument.op)
            }
        }

        private void getTableHeader_Click(object sender, EventArgs e)
        {
            string fileName = (@"D:\OfficeDev\Excel\201602\Test.xlsx");
            var items =ExcelTableHeader.GetTableHeader(fileName, "Sheet1");
            string line = string.Join(",", items.ToArray());
            MessageBox.Show(line);

        }

        private void PPTInsertShape_Click(object sender, EventArgs e)
        {
            PPTInsertShape ppt = new Class.OpenXmlHelper.PPTHelper.PPTInsertShape();
            ppt.InserShape(@"D:\OfficeDev\PPT\201608\InsertShape - Copy.pptx");
            MessageBox.Show("ok");
        }

        private void CoreFileProperties_Click(object sender, EventArgs e)
        {
            CoreFilePropertiesClass c = new CoreFilePropertiesClass();
            c.ChangeCoreFileProperties(@"D:\OfficeDev\OpenXML\Excel\CreateTime.xlsx");
        }

        private void ImageATL_Click(object sender, EventArgs e)
        {
            OfficeAPI.Class.OpenXmlHelper.ExcelHelper.ImageATL imageATL = new ImageATL();
            imageATL.ChangeImageATL(@"D:\OfficeDev\OpenXML\Excel\ImageAlt.xlsx");
        }

        private void InsertImg_Click(object sender, EventArgs e)
        {
            InsertImgToMaster iit = new Class.OpenXmlHelper.PPTHelper.InsertImgToMaster();
            iit.InsertImg(@"D:\OfficeDev\OpenXML\PPT\201608\InsertImg - Copy.pptx");
            MessageBox.Show("ok");
        }

        private void ExcelColumn_Click(object sender, EventArgs e)
        {
            ColumnWidth cw = new ColumnWidth();
            cw.getDColumnWidth(@"D:\OfficeDev\OpenXML\Excel\20160822\General Tool 0806cT.xlsm");
            //cw.getWidth(@"D:\OfficeDev\OpenXML\Excel\20160822\General Tool 0806c.xlsm");
        }

        private void SMTPbtn_Click(object sender, EventArgs e)
        {
            SmtpClient client = new SmtpClient("smtp.office365.com", 587);
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential("v-tazho@msdnofficedevteam.onmicrosoft.com", "Edward1121");
            MailAddress from = new MailAddress("v-tazho@msdnofficedevteam.onmicrosoft.com", String.Empty, System.Text.Encoding.UTF8);
            MailAddress to = new MailAddress("v-tazho@hotmail.com");
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(from, to);
            message.Body = "The message I want to send.";
            client.Send(message);
            MessageBox.Show("ok");
        }

        private void ChangeField_Click(object sender, EventArgs e)
        {
            ChangeFieldClass cfc = new ChangeFieldClass();
            cfc.GetSetField(@"D:\OfficeDev\OpenXML\Word\201609\Input.docm");
        }

        private void OneNotePage_Click(object sender, EventArgs e)
        {
            OneNoteExtension oneNote = new OneNoteExtension();
            oneNote.OneNotePage();
            MessageBox.Show("ok");
        }

        private void RequestAPI_Click(object sender, EventArgs e)
        {
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create("http://10.168.172.127/");

            request.Method = "GET";
            request.UseDefaultCredentials = true;
            request.PreAuthenticate = true;
            //request.Credentials = new NetworkCredential("v-tazho@microsoft.com", "Edward1001", "fareast");

            HttpWebResponse response = (HttpWebResponse)request.GetResponse(); // Raises Unauthorized Exception

            
        }

        private void OneNoteCreatePage_Click(object sender, EventArgs e)
        {
            OneNoteExtension oneNote = new OneNoteExtension();
            oneNote.OneNoteNewPage();
            MessageBox.Show("ok");
             
        }

        private void FileCopyExcel_Click(object sender, EventArgs e)
        {
            File.Copy(@"D:\OfficeDev\Excel\201611\Test.xlsx", @"D:\OfficeDev\Excel\201611\New0.xlsx", true);
        }

        private void CopyWorkbook_Click(object sender, EventArgs e)
        {

        }

        private void setFont_Click(object sender, EventArgs e)
        {
            try
            {
                _word.Application winword = new _word.Application();
                winword.ShowAnimation = false;
                winword.Visible = false;

                object missing = System.Reflection.Missing.Value;

                _word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                _word.Paragraph pp = document.Paragraphs.Add(missing);
                pp.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                pp.Range.Font.Name = "B Nazanin";
                pp.Range.Font.Size = 16;
                pp.Range.Text = "امید منصورنیا";
                pp.Range.InsertParagraphAfter();


                object filename = @"D:\OfficeDev\Word\201612\temp1.docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                MessageBox.Show("Document created successfully !");

                System.Diagnostics.Process.Start("WINWORD.EXE", filename.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        // Signals when the resolve has finished.
        public static ManualResetEvent GetHostEntryFinished =
            new ManualResetEvent(false);
        private void DnsMethod_Click(object sender, EventArgs e)
        {
            //DoGetHostEntryAsync("vdi-v-tazho");
            DoGetHostEntryAsync("10.168.197.122");
        }
        // Determine the Internet Protocol (IP) addresses for 
        // this host asynchronously.
        public static void DoGetHostEntryAsync(string hostname)
        {
            GetHostEntryFinished.Reset();
            ResolveState ioContext = new ResolveState(hostname);

            Dns.BeginGetHostEntry(ioContext.host,
                new AsyncCallback(GetHostEntryCallback), ioContext);

            // Wait here until the resolve completes (the callback 
            // calls .Set())
            GetHostEntryFinished.WaitOne();

            Console.WriteLine("EndGetHostEntry({0}) returns:", ioContext.host);

            foreach (IPAddress ip in ioContext.IPs.AddressList)
            {
                Console.WriteLine("    {0}", ip);
            }

        }
        // Record the IPs in the state object for later use.
        public static void GetHostEntryCallback(IAsyncResult ar)
        {
            ResolveState ioContext = (ResolveState)ar.AsyncState;

            ioContext.IPs = Dns.EndGetHostEntry(ar);
            GetHostEntryFinished.Set();
        }

        private void DoGetHostAddresses_Click(object sender, EventArgs e)
        {
            IPAddress[] ips;
            //ips = Dns.GetHostAddresses("ip address"); 
            ips = Dns.GetHostAddresses("");         
            foreach (IPAddress ip in ips)
            {
                Console.WriteLine("    {0}", ip);
            }

        }       

    }
    // Define the state object for the callback. 
    // Use hostName to correlate calls with the proper result.
    public class ResolveState
    {
        string hostName;
        IPHostEntry resolvedIPs;

        public ResolveState(string host)
        {
            hostName = host;
        }

        public IPHostEntry IPs
        {
            get { return resolvedIPs; } 
            set {resolvedIPs = value;}}
        public string host {get {return hostName;} 
            set {hostName = value;}}
    }

    public static class Extensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
                return xdoc;
            using (StreamReader streamReader = new StreamReader(part.GetStream()))
                xdoc = XDocument.Load(XmlReader.Create(streamReader));
            part.AddAnnotation(xdoc);
            return xdoc;
        }

        public static string StringConcatenate(this IEnumerable<string> source)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
                sb.Append(s);
            return sb.ToString();
        }
        public static void PutXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.GetXDocument();
            if (xdoc != null)
            {
                // Serialize the XDocument object back to the package.
                using (XmlWriter xw =
                    XmlWriter.Create(part.GetStream
                   (FileMode.Create, FileAccess.Write)))
                {
                    xdoc.Save(xw);
                }
            }
        }
    }
    public static class W
    {
        public static XNamespace w =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static XName t = w + "t";
        public static XName r = w + "r";
        public static XName del = w + "del";
        public static XName body = w + "body";
        public static XName p = w + "p";
        public static XName moveFrom = w + "moveFrom";
    }

}
