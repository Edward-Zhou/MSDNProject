using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    public class RtfImport
    {
        public static void ImportRtf(string FileName)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(FileName, true))
              {
                string altChunkId = "AltChunkId5";

                MainDocumentPart mainDocPart = doc.MainDocumentPart;
                AlternativeFormatImportPart chunk = mainDocPart.AddAlternativeFormatImportPart(
                    AlternativeFormatImportPartType.Rtf, altChunkId);
                //set rtfEncodedString with rtf string
                string rtfEncodedString = @"{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard This is some {\b bold} text.\par}";

                using (MemoryStream ms = new MemoryStream(Encoding.ASCII.GetBytes(rtfEncodedString)))
                {
                  chunk.FeedData(ms);
                }

                AltChunk altChunk = new AltChunk();
                altChunk.Id = altChunkId;

                mainDocPart.Document.Body.InsertAfter(
                  altChunk, mainDocPart.Document.Body.Elements<Paragraph>().Last());

                mainDocPart.Document.Save();

              }
        }
    }
}
