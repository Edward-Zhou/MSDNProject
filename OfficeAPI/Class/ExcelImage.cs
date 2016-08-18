using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRW = System.Drawing;
using IMG = System.Drawing.Imaging; 

namespace OfficeAPI.Class
{
    class ExcelImage
    {
        private static System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private static SpreadsheetDocument excelDocument;
        public static void CreateSummaryExcelDoc(string fileName)
        {           

            //Create excel document
            using ( excelDocument =
                SpreadsheetDocument.Open(fileName, true))
            {

                var workbookPart = excelDocument.WorkbookPart ;
                excelDocument.WorkbookPart.Workbook = new Workbook();
                excelDocument.WorkbookPart.Workbook.Sheets = new Sheets();

                var sheetPart = excelDocument.WorkbookPart.WorksheetParts.FirstOrDefault(); //excelDocument.WorkbookPart.AddNewPart<WorksheetPart>();

                SheetData sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);
                Worksheet ws = sheetPart.Worksheet;


                //Generate header image parts.
                //VmlDrawingPart vmlDrawingPart1 = sheetPart.AddNewPart<VmlDrawingPart>("rId2");
                BuildUriPartDictionary();
                VmlDrawingPart vmlDrawingPart1 = UriPartDictionary["/xl/worksheets/sheet1.xml"].AddNewPart<VmlDrawingPart>("rId2");
                GenerateVmlDrawingPart1Content(vmlDrawingPart1);

                ImagePart imagePart1 = vmlDrawingPart1.AddNewPart<ImagePart>("image/png", "rId1");
                GenerateImagePart1Content(imagePart1);

                ChangeWorksheetPart1((WorksheetPart)UriPartDictionary["/xl/worksheets/sheet1.xml"]);


                Sheets sheets = excelDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                string relationshipId = excelDocument.WorkbookPart.GetIdOfPart(sheetPart);
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = 1, Name = "Projects" };
                sheets.Append(sheet);


                WorkbookStylesPart stylesPart = excelDocument.WorkbookPart.WorkbookStylesPart;
                stylesPart.Stylesheet = GenerateStyleSheet();
                stylesPart.Stylesheet.Save();


            }
        }
        private static void BuildUriPartDictionary()
        {
            System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
            queue.Enqueue(excelDocument);
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
        private static void GenerateVmlDrawingPart1Content(VmlDrawingPart vmlDrawingPart1)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"LH\" o:spid=\"_x0000_s1025\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:168pt;height:44.25pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"contisystems.png\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }

        private static void GenerateImagePart1Content(ImagePart imagePart1)
        {
            DRW.Image image = DRW.Image.FromFile(@"C:\Users\Administrator\Desktop\1.png");
            using (MemoryStream stream = new MemoryStream())
            {
                // Save image to stream.
                image.Save(stream, IMG.ImageFormat.Png);
                string imagePart1Data = Convert.ToBase64String(stream.ToArray());
                System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
                imagePart1.FeedData(data);
                data.Close();
            }

        }
        private static void ChangeWorksheetPart1(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = worksheetPart1.Worksheet;

            HeaderFooter headerFooter1 = new HeaderFooter();
            FirstHeader firstHeader1 = new FirstHeader();
            firstHeader1.Text = "&L&G";

            headerFooter1.Append(firstHeader1);
            worksheet1.Append(headerFooter1);

            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter1 = new LegacyDrawingHeaderFooter() { Id = "rId2" };
            worksheet1.Append(legacyDrawingHeaderFooter1);
        }
        private static Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                    new Font(                                                               // Index 0 - The default font.
                        new FontSize() { Val = 9 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 1 - The bold font.
                        new Bold(),
                        new FontSize() { Val = 10 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 2 - The Italic font.
                        new Italic(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 2 - The Times Roman font. with 16 size
                        new FontSize() { Val = 16 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" })
                ),
                new Fills(
                    new DocumentFormat.OpenXml.Spreadsheet.Fill(                                                           // Index 0 - The default fill.
                        new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.None }),
                    new DocumentFormat.OpenXml.Spreadsheet.Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                        new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.Gray125 }),
                    new DocumentFormat.OpenXml.Spreadsheet.Fill(                                                           // Index 2 - The Snow 3 Pastel fill.
                        new DocumentFormat.OpenXml.Spreadsheet.PatternFill(
                            new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor() { Rgb = new HexBinaryValue() { Value = "cdc9c9" } }
                        ) { PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(                                                         // Index 0 - The default border.
                        new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(),
                        new DocumentFormat.OpenXml.Spreadsheet.RightBorder(),
                        new DocumentFormat.OpenXml.Spreadsheet.TopBorder(),
                        new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(),
                        new DiagonalBorder()),
                    new Border(                                                         // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                        new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(
                            new Color() { Auto = true }
                        ) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.RightBorder(
                            new Color() { Auto = true }
                        ) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.TopBorder(
                            new Color() { Auto = true }
                        ) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(
                            new Color() { Auto = true }
                        ) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                // Diferentes StyleIndex's para utilizar
                new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1 },                         // Index 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat() { FontId = 1, FillId = 2, BorderId = 1, ApplyFont = true },       // Index 1 - Header Bold and Filled 
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 - Italic
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 1, ApplyFont = true },       // Index 3 - Times Roman
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Index 4 - Snow 3 Fill
                    new CellFormat(                                                                    // Index 5 - Alignment
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    ) { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }      // Index 6 - Borders
                )
            ); // return
        }
        private static System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }
    }
}
