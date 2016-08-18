using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    class DeleteContentControl
    {
        WordprocessingDocument WordDoc;
        public void RemoveContentControl(string filePath)
        {
            using (WordDoc = WordprocessingDocument.Open(filePath,true))
            {
                MainDocumentPart main = WordDoc.MainDocumentPart;
                SdtBlock[] sdtBlock = main.Document.Body.Descendants<SdtBlock>().ToArray();
                foreach (SdtBlock sdt in sdtBlock)
                {
                    sdt.Remove();
                }
            }
        }
    }
}
