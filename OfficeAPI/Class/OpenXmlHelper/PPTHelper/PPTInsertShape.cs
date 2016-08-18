using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;


namespace OfficeAPI.Class.OpenXmlHelper.PPTHelper
{
    public  class PPTInsertShape
    {
        public void InserShape(string filePath)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(filePath, true))
            {
                PresentationPart pptPart = ppt.PresentationPart;
                SlidePart slidePart = (SlidePart)pptPart.SlideParts.FirstOrDefault();
                ChangeSlidePart(slidePart);
            }
        }
        public void ChangeSlidePart(SlidePart slidePart1)
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            Shape shape1 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "TextBox 3" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 4549966L, Y = 1927952L };
            A.Extents extents1 = new A.Extents() { Cx = 1795750L, Cy = 369332L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill1 = new A.NoFill();

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);

            TextBody textBody1 = new TextBody();

            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            bodyProperties1.Append(shapeAutoFit1);
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "zh-CN", Dirty = false };
            runProperties1.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text1 = new A.Text();
            text1.Text = "Test";

            run1.Append(runProperties1);
            run1.Append(text1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "zh-CN", AlternativeLanguage = "en-US", Dirty = false };

            paragraph1.Append(run1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);
            shapeTree1.Append(shape1);
        }
    }
}
