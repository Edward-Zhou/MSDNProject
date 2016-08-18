using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Windows.Forms;

namespace OfficeAPI.Class.OpenXmlHelper.ExcelHelper
{
    public class ImageATL
    {
        public void ChangeImageATL(string filePath)
        { 
            using(SpreadsheetDocument sheetDocument=SpreadsheetDocument.Open(filePath,true))
            {
                WorkbookPart workbookPart = sheetDocument.WorkbookPart;
                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                {
                    DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
                    ChangeDrawingsPart(drawingsPart);
                }
            }
        }
        public void ChangeDrawingsPart(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = drawingsPart1.WorksheetDrawing;
            foreach(Xdr.TwoCellAnchor twoCellAnchor in worksheetDrawing1.Elements<Xdr.TwoCellAnchor>())
            {
                Xdr.Picture picture1 = twoCellAnchor.GetFirstChild<Xdr.Picture>();
                Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = picture1.GetFirstChild<Xdr.NonVisualPictureProperties>();
                Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = nonVisualPictureProperties1.GetFirstChild<Xdr.NonVisualDrawingProperties>();
                MessageBox.Show(nonVisualDrawingProperties1.Description + "; " + nonVisualDrawingProperties1.Title);
            }

            //Xdr.TwoCellAnchor twoCellAnchor1 = worksheetDrawing1.GetFirstChild<Xdr.TwoCellAnchor>();
            //Xdr.TwoCellAnchor twoCellAnchor2 = worksheetDrawing1.Elements<Xdr.TwoCellAnchor>().ElementAt(1);

            //Xdr.Picture picture1 = twoCellAnchor1.GetFirstChild<Xdr.Picture>();

            //Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = picture1.GetFirstChild<Xdr.NonVisualPictureProperties>();

            //Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = nonVisualPictureProperties1.GetFirstChild<Xdr.NonVisualDrawingProperties>();
            //nonVisualDrawingProperties1.Description = "ALTDescription1";
            //nonVisualDrawingProperties1.Title = "ALTTitle1";

            //Xdr.Picture picture2 = twoCellAnchor2.GetFirstChild<Xdr.Picture>();

            //Xdr.NonVisualPictureProperties nonVisualPictureProperties2 = picture2.GetFirstChild<Xdr.NonVisualPictureProperties>();

            //Xdr.NonVisualDrawingProperties nonVisualDrawingProperties2 = nonVisualPictureProperties2.GetFirstChild<Xdr.NonVisualDrawingProperties>();
            //nonVisualDrawingProperties2.Description = "ALTDescription2";
            //nonVisualDrawingProperties2.Title = "ALTTitle2";
        }

    }
}
