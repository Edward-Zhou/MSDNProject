using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAPI.Class.OpenXmlHelper
{
    class ExcelTableHeader
    {
        public static List<string> GetTableHeader(string FileName, string worksheetName)
        {
            List<string> items = new List<string>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(FileName, true))
            {

                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
                if (sheets.Count() == 0)
                {
                    return null;
                }

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

                //get all the tables
                foreach (TableDefinitionPart tablesPart in worksheetPart.TableDefinitionParts)
                {

                    Table table = tablesPart.Table;
                    // get  all the table columns
                    foreach (TableColumn tableColunm in tablesPart.Table.TableColumns)
                    {
                        items.Add(tableColunm.Name.Value);
                    }
                }
                return items;
            }
        }

    }
}
