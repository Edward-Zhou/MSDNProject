using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAPI.Class.OpenXmlHelper.PPTHelper
{
    class AddNoteClass

    {
        public static void AddNote(string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                string relId = (slideIds[index] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                // Build a StringBuilder object.
                StringBuilder paragraphText = new StringBuilder();

                // Get the inner text of the slide:
                IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();

                NotesSlidePart notesSlidePart1 = slide.AddNewPart<NotesSlidePart>();
                NotesSlide notesSlide = new NotesSlide(
                new CommonSlideData(new ShapeTree(
                  new P.NonVisualGroupShapeProperties(
                  new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                  new P.NonVisualGroupShapeDrawingProperties(),
                  new ApplicationNonVisualDrawingProperties()),
                  new GroupShapeProperties(new A.TransformGroup()),
                  new P.Shape(
                  new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                  new P.ShapeProperties(),
                  new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.EndParagraphRunProperties()))))),
                new ColorMapOverride(new A.MasterColorMapping()));
                notesSlidePart1.NotesSlide = notesSlide;
                ChangePresentationPart(part);

            }
        }
        private static NotesSlidePart CreateNotesSlidePart(SlidePart slidePart1)
        {
            NotesSlidePart notesSlidePart1 = slidePart1.AddNewPart<NotesSlidePart>("rId3");
            NotesSlide notesSlide = new NotesSlide(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new A.TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties()))))),
            new ColorMapOverride(new A.MasterColorMapping()));
            notesSlidePart1.NotesSlide = notesSlide;
            return notesSlidePart1;
        }

        private  void ChangeNotesSlidePart1(NotesSlidePart notesSlidePart1)
        {
            NotesSlide notesSlide1 = notesSlidePart1.NotesSlide;

            CommonSlideData commonSlideData1 = notesSlide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            Shape shape1 = shapeTree1.Elements<Shape>().ElementAt(1);

            TextBody textBody1 = shape1.GetFirstChild<TextBody>();

            A.Paragraph paragraph1 = textBody1.GetFirstChild<A.Paragraph>();

            A.EndParagraphRunProperties endParagraphRunProperties1 = paragraph1.GetFirstChild<A.EndParagraphRunProperties>();

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "zh-CN" };
            runProperties1.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text1 = new A.Text();
            text1.Text = "Test";

            run1.Append(runProperties1);
            run1.Append(text1);
            paragraph1.InsertBefore(run1, endParagraphRunProperties1);
            endParagraphRunProperties1.Dirty = null;
        }

        public static void ChangePresentationPart(PresentationPart presentationPart1)
        {
            Presentation presentation1 = presentationPart1.Presentation;

            SlideIdList slideIdList1 = presentation1.GetFirstChild<SlideIdList>();

            NotesMasterIdList notesMasterIdList1 = new NotesMasterIdList();
            NotesMasterId notesMasterId1 = new NotesMasterId() { Id = "rId3" };

            notesMasterIdList1.Append(notesMasterId1);
            presentation1.InsertBefore(notesMasterIdList1, slideIdList1);
        }



    }
}
