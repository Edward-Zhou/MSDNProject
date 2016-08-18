using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeAPI.Class.OpenXmlHelper
{
    class SetPPTShapeColorClass
    {
        public static void SetPPTShapeColor(string docName)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[0] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);
                //ChangeSlidePart(slide);
                if (slide != null)
                {
                    // Get the shape tree that contains the shape to change.
                    ShapeTree tree = slide.Slide.CommonSlideData.ShapeTree;
                    foreach (Shape s in tree.Elements<Shape>())
                    {
                        string ShapeText = s.InnerText;
                        if (ShapeText == "Test1")
                        {
                            ShapeProperties shapeProperties1 = s.GetFirstChild<ShapeProperties>();
                            ShapeStyle shapeStyle1 = s.GetFirstChild<ShapeStyle>();
                            A.SolidFill solidFill1 = shapeProperties1.GetFirstChild<A.SolidFill>();
                            if (solidFill1 != null)
                            {
                                //change solid fill color
                                A.SchemeColor schemeColor1 = solidFill1.GetFirstChild<A.SchemeColor>();
                                schemeColor1.Val = A.SchemeColorValues.Accent6;
                            }
                            //change fill reference color
                            A.FillReference fillReference1 = shapeStyle1.GetFirstChild<A.FillReference>();
                            A.SchemeColor schemeColor2 = fillReference1.GetFirstChild<A.SchemeColor>();
                            schemeColor2.Val = A.SchemeColorValues.Accent1;
                        }
                    }
                    // Get the first shape in the shape tree.
                    Shape shape = tree.GetFirstChild<Shape>();

                    if (shape != null)
                    {
                        // Get the style of the shape.
                        ShapeStyle style = shape.ShapeStyle;

                        // Get the fill reference.
                        Drawing.FillReference fillRef = style.FillReference;
                        //Drawing.SolidFill solidFill=style.
                        // Set the fill color to SchemeColor Accent 6;
                        fillRef.SchemeColor = new Drawing.SchemeColor();
                        fillRef.SchemeColor.Val = Drawing.SchemeColorValues.Accent6;
                        // Save the modified slide.
                        slide.Slide.Save();
                    }
                }
            }
        }
        public static void ChangeSlidePart(SlidePart slidePart1)
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            Shape shape1 = shapeTree1.GetFirstChild<Shape>();

            ShapeProperties shapeProperties1 = shape1.GetFirstChild<ShapeProperties>();
            ShapeStyle shapeStyle1 = shape1.GetFirstChild<ShapeStyle>();

            A.SolidFill solidFill1 = shapeProperties1.GetFirstChild<A.SolidFill>();

            A.SchemeColor schemeColor1 = solidFill1.GetFirstChild<A.SchemeColor>();
            schemeColor1.Val = A.SchemeColorValues.Accent6;

            A.FillReference fillReference1 = shapeStyle1.GetFirstChild<A.FillReference>();

            A.SchemeColor schemeColor2 = fillReference1.GetFirstChild<A.SchemeColor>();
            schemeColor2.Val = A.SchemeColorValues.Accent1;
        }
    }
}
