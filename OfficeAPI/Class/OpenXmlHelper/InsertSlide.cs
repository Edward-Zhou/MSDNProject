using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeAPI.Class.OpenXmlHelper
{
    class InsertSlide
    {
        public static void InsertNewSlide(string presentationFile, int position, string layoutName)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {

                InsertNewSlide(presentationDocument, position, layoutName);
                
            }
        }

        public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string layoutName)
        {
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            OpenXML.Slide slide = new OpenXML.Slide(new CommonSlideData(new ShapeTree()));

            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            slide.Save(slidePart);

            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();

            SlideLayoutPart slideLayoutPart = slideMasterPart.SlideLayoutParts.SingleOrDefault(sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName, StringComparison.OrdinalIgnoreCase));

            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);

            slidePart.Slide.CommonSlideData = (CommonSlideData)slideMasterPart.SlideLayoutParts.SingleOrDefault(sl => sl.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName)).SlideLayout.CommonSlideData.Clone();

            using (Stream stream = slideLayoutPart.GetStream())
            {
                slidePart.SlideLayoutPart.FeedData(stream);

            }

            foreach (ImagePart iPart in slideLayoutPart.ImageParts)
            {
                ImagePart newImagePart = slidePart.AddImagePart(iPart.ContentType, slideLayoutPart.GetIdOfPart(iPart));
                newImagePart.FeedData(iPart.GetStream());
            }
            ChangeSlidePart(slidePart);
            uint maxSlideId = 1;
            SlideId prevSlideId = null;
            var slideIdList = presentationPart.Presentation.SlideIdList;
            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }
            maxSlideId++;
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            presentationPart.Presentation.Save();
        }
        public static void ChangeSlidePart(SlidePart slidePart1)
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            Picture picture1 = shapeTree1.GetFirstChild<Picture>();

            picture1.Remove();
        }
    }
    
}
