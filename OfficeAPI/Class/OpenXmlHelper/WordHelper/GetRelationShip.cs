using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAPI.Class.OpenXmlHelper.WordHelper
{
    class GetRelationShip
    {
        WordprocessingDocument WordDoc;
        public void getExternalRelationShip(string filePath)
        {
            Package p = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            
            foreach (PackageRelationship pr in p.GetRelationships())
            {
                
                string id = pr.Id;
            }
            //ExternalRelationship pPart = mPart.GetExternalRelationship("rId4");
            //using (WordDoc=WordprocessingDocument.Open(filePath,false))
            //{
            //    MainDocumentPart mPart = WordDoc.MainDocumentPart;



                
            //}
        }
    }
}
