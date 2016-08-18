using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cdr = DocumentFormat.OpenXml.Drawing.ChartDrawing;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeAPI.Class.ExcelAutomation
{
    class AddShapeToChartClass
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private SpreadsheetDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = SpreadsheetDocument.Open(filePath, true))
            {
                ChangeParts();
            }
        }

        private void ChangeParts()
        {
            //Stores the referrences to all the parts in a dictionary.
            BuildUriPartDictionary();
            //Adds new parts or new relationships.
            AddParts();
            //Changes the contents of the specified parts.
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeChartPart1(((ChartPart)UriPartDictionary["/xl/charts/chart1.xml"]));
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
        /// Adds new parts or new relationship between parts.
        /// </summary>
        private void AddParts()
        {
            //Generate new parts.
            ChartDrawingPart chartDrawingPart1 = UriPartDictionary["/xl/charts/chart1.xml"].AddNewPart<ChartDrawingPart>("rId3");
            GenerateChartDrawingPart1Content(chartDrawingPart1);

        }

        private void GenerateChartDrawingPart1Content(ChartDrawingPart chartDrawingPart1)
        {
            C.UserShapes userShapes1 = new C.UserShapes();
            userShapes1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            Cdr.RelativeAnchorSize relativeAnchorSize1 = new Cdr.RelativeAnchorSize();
            relativeAnchorSize1.AddNamespaceDeclaration("cdr", "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing");

            Cdr.FromAnchor fromAnchor1 = new Cdr.FromAnchor();
            Cdr.XPosition xPosition1 = new Cdr.XPosition();
            xPosition1.Text = "0.32917";
            Cdr.YPosition yPosition1 = new Cdr.YPosition();
            yPosition1.Text = "0.23438";

            fromAnchor1.Append(xPosition1);
            fromAnchor1.Append(yPosition1);

            Cdr.ToAnchor toAnchor1 = new Cdr.ToAnchor();
            Cdr.XPosition xPosition2 = new Cdr.XPosition();
            xPosition2.Text = "0.57708";
            Cdr.YPosition yPosition2 = new Cdr.YPosition();
            yPosition2.Text = "0.39757";

            toAnchor1.Append(xPosition2);
            toAnchor1.Append(yPosition2);

            Cdr.Shape shape1 = new Cdr.Shape() { Macro = "", TextLink = "" };

            Cdr.NonVisualShapeProperties nonVisualShapeProperties1 = new Cdr.NonVisualShapeProperties();
            Cdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Cdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Rounded Rectangle 1" };
            Cdr.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new Cdr.NonVisualShapeDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);

            Cdr.ShapeProperties shapeProperties1 = new Cdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            transform2D1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.Offset offset1 = new A.Offset() { X = 1504950L, Y = 642938L };
            A.Extents extents1 = new A.Extents() { Cx = 1133475L, Cy = 447675L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RoundRectangle };
            presetGeometry1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            Cdr.Style style1 = new Cdr.Style();

            A.LineReference lineReference1 = new A.LineReference() { Index = (UInt32Value)2U };
            lineReference1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade1 = new A.Shade() { Val = 50000 };

            schemeColor1.Append(shade1);

            lineReference1.Append(schemeColor1);

            A.FillReference fillReference1 = new A.FillReference() { Index = (UInt32Value)1U };
            fillReference1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor2);

            A.EffectReference effectReference1 = new A.EffectReference() { Index = (UInt32Value)0U };
            effectReference1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor3);

            A.FontReference fontReference1 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            fontReference1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference1.Append(schemeColor4);

            style1.Append(lineReference1);
            style1.Append(fillReference1);
            style1.Append(effectReference1);
            style1.Append(fontReference1);

            Cdr.TextBody textBody1 = new Cdr.TextBody();

            A.BodyProperties bodyProperties1 = new A.BodyProperties() { VerticalOverflow = A.TextVerticalOverflowValues.Clip };
            bodyProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ListStyle listStyle1 = new A.ListStyle();
            listStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.Paragraph paragraph1 = new A.Paragraph();
            paragraph1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "zh-CN" };

            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(style1);
            shape1.Append(textBody1);

            relativeAnchorSize1.Append(fromAnchor1);
            relativeAnchorSize1.Append(toAnchor1);
            relativeAnchorSize1.Append(shape1);

            userShapes1.Append(relativeAnchorSize1);

            chartDrawingPart1.UserShapes = userShapes1;
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2016-05-23T02:45:31Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeChartPart1(ChartPart chartPart1)
        {
            C.ChartSpace chartSpace1 = chartPart1.ChartSpace;

            C.Chart chart1 = chartSpace1.GetFirstChild<C.Chart>();

            C.PlotArea plotArea1 = chart1.GetFirstChild<C.PlotArea>();

            C.BarChart barChart1 = plotArea1.GetFirstChild<C.BarChart>();
            C.CategoryAxis categoryAxis1 = plotArea1.GetFirstChild<C.CategoryAxis>();
            C.ValueAxis valueAxis1 = plotArea1.GetFirstChild<C.ValueAxis>();

            C.AxisId axisId1 = barChart1.GetFirstChild<C.AxisId>();
            C.AxisId axisId2 = barChart1.Elements<C.AxisId>().ElementAt(1);
            axisId1.Val = (UInt32Value)895575552U;
            axisId2.Val = (UInt32Value)895575944U;

            C.AxisId axisId3 = categoryAxis1.GetFirstChild<C.AxisId>();
            C.CrossingAxis crossingAxis1 = categoryAxis1.GetFirstChild<C.CrossingAxis>();
            axisId3.Val = (UInt32Value)895575552U;
            crossingAxis1.Val = (UInt32Value)895575944U;

            C.AxisId axisId4 = valueAxis1.GetFirstChild<C.AxisId>();
            C.CrossingAxis crossingAxis2 = valueAxis1.GetFirstChild<C.CrossingAxis>();
            axisId4.Val = (UInt32Value)895575944U;
            crossingAxis2.Val = (UInt32Value)895575552U;

            C.UserShapesReference userShapesReference1 = new C.UserShapesReference() { Id = "rId3" };
            chartSpace1.Append(userShapesReference1);
        }

    }
}
