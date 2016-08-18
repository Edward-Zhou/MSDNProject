using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeAPI.Class.OpenXmlHelper
{
    public class CoreFilePropertiesClass
    {
        public void ChangeCoreFileProperties(string filePath)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                CoreFilePropertiesPart cp = spreadsheetDocument.CoreFilePropertiesPart;
                ChangeCoreFilePropertiesPart(cp);
            }
        }
        public void ChangeCoreFilePropertiesPart(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Creator = "";
            package.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2016-08-05T08:29:39Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2016-08-05T08:29:39Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            package.PackageProperties.LastModifiedBy = "";
        }

    }
}
