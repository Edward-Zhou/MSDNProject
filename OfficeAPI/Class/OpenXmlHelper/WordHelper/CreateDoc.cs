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
    class CreateDoc
    {
        public void CreateDocument()
        {
            string sFile = @"D:\OfficeDev\OpenXML\Word\Hello World.dotx";
            if (File.Exists(sFile.Replace("dotx", "docx")))
                File.Delete(sFile.Replace("dotx", "docx"));               
            File.Copy(sFile, sFile.Replace("dotx", "docx"));
            WordprocessingDocument moWordDoc;
            //create word doc
            moWordDoc = WordprocessingDocument.Open(sFile.Replace("dotx", "docx"), true);
            //change doc type, which ensures the document is no longer marked as a Template but rather as a Document
            moWordDoc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            AttachedTemplate oAttachedTemplate = new AttachedTemplate();
            oAttachedTemplate.Id = "relationId1";
            //append the AttachedTemplate to the DocumentSettingsPart 
            MainDocumentPart oMainPart = moWordDoc.MainDocumentPart;
            DocumentSettingsPart oDocSettings = oMainPart.DocumentSettingsPart;
            oDocSettings.Settings.Append(oAttachedTemplate);
            oDocSettings.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", new Uri(sFile, UriKind.Absolute), "relationId1");

        }
        public void CreateDocFromTem()
        {
            string sFile = @"D:\OfficeDev\OpenXML\Word\Hello World.dotx";
            if (File.Exists(sFile.Replace("dotx", "docx")))
                File.Delete(sFile.Replace("dotx", "docx"));
            File.Copy(sFile, sFile.Replace("dotx", "docx"));
            UnicodeEncoding uniEncoding = new UnicodeEncoding();
            FileStream fs = new FileStream(sFile, FileMode.Open, FileAccess.Read);
            MemoryStream templateStream=new MemoryStream();
            fs.CopyTo(templateStream);
            using (MemoryStream documentStream = new MemoryStream((int)templateStream.Length))
            {
                templateStream.Position = 0L;
                byte[] buffer = new byte[2048];
                int length = buffer.Length;
                int size;
                while ((size = templateStream.Read(buffer, 0, length)) != 0)
                {
                    documentStream.Write(buffer, 0, size);
                }
                documentStream.Position = 0L;

                // Modify the document to ensure it is correctly marked as a Document and not Template
                using (WordprocessingDocument document = WordprocessingDocument.Open(documentStream, true))
                {
                    document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                }

                // Add the XML into the document and save to the correct location.               
                File.WriteAllBytes(sFile.Replace("dotx", "docx"), documentStream.ToArray());
            }
        }
        public void CreateDocFromTembak()
        {
            string sFile = @"D:\OfficeDev\OpenXML\Word\Hello World.dotx";
            if (File.Exists(sFile.Replace("dotx", "docx")))
                File.Delete(sFile.Replace("dotx", "docx"));
            File.Copy(sFile, sFile.Replace("dotx", "docx"));
            UnicodeEncoding uniEncoding = new UnicodeEncoding();
            FileStream fs = new FileStream(sFile, FileMode.Open, FileAccess.Read);
            MemoryStream templateStream = new MemoryStream();
            fs.CopyTo(templateStream);
            using (MemoryStream documentStream = new MemoryStream((int)templateStream.Length))
            {

                // Modify the document to ensure it is correctly marked as a Document and not Template
                using (WordprocessingDocument document = WordprocessingDocument.Open(documentStream, true))
                {
                    document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                }

                // Add the XML into the document and save to the correct location.               
                File.WriteAllBytes(sFile.Replace("dotx", "docx"), documentStream.ToArray());
            }
        }

    }
}
