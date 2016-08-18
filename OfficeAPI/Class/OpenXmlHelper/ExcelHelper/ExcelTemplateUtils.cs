using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
//using Services.ViewModels;
using System.Threading.Tasks;
using OpenXml.Excel;

namespace Excel.Utilities
{
    public class ExcelTemplateUtils
    {
        public static bool TemplateDownloadStuffWithData()
        {
            List<int> list = new List<int> { 1,2,3};
            bool result = false;
            if (list.Count > 0)
            {
                using (ExcelXmlMapper msExcel = new ExcelXmlMapper())
                {
                    uint styleValue = 0;
                    //msExcel.Open(@"D:\Edward\Project\MSDNProject\MSDNProject\OfficeAPI\Class\OpenXmlHelper\ExcelHelper\NewSpreadsheet3.xlsx", true);
                    msExcel.InsertText(@"D:\Edward\Project\MSDNProject\MSDNProject\OfficeAPI\Class\OpenXmlHelper\ExcelHelper\NewSpreadsheet4.xlsx", "Hellow word");
                    //// fill in the identifiers we need for when we read the file
                    //msExcel.SetCurrentSheet("Dropdown");
                    //msExcel.SetCellValue("AA6", "1", styleValue, Celltype.isNumber);
                    ////msExcel.SetCellValue("O4", "2", styleValue, Celltype.isNumber);
                    ////msExcel.SetCellValue("O5", "3", styleValue, Celltype.isNumber);
                    //msExcel.SaveCurrentSheet();

                    //// fill in the consumption data
                    //msExcel.SetCurrentSheet("Data");
                    //int nrow = 6;
                    //foreach (var ocData in list)
                    //{
                    //    // common data in each sheet
                    //    //msExcel.SetCellValue("AA" + nrow.ToString(), "Some String", styleValue, Celltype.isString);
                    //    //msExcel.SetCellValue("B" + nrow.ToString(), "Another String", styleValue, Celltype.isString);
                    //    //msExcel.SetCellValue("C" + nrow.ToString(), "String Value", styleValue, Celltype.isString);

                    //    // save the id of the record for when we read the data back
                    //    //msExcel.SetCellValue("AA" + nrow.ToString(), "21", styleValue, Celltype.isNumber);
                    //    //msExcel.SetCellValue("AB" + nrow.ToString(), "22", styleValue, Celltype.isNumber);
                    //    //msExcel.SetCellValue("AC" + nrow.ToString(), "23", styleValue, Celltype.isNumber);
                    //    //msExcel.SetCellValue("AD" + nrow.ToString(), "24", styleValue, Celltype.isNumber);

                    //    nrow++;
                    //}

                    //// fill in the identifiers we need for when we read the file
                    //msExcel.SetCurrentSheet("Info");
                    //msExcel.SetCellValue("C8", "A Name", styleValue, Celltype.isString);
                    //msExcel.SetCellValue("C9", "Another Name", styleValue, Celltype.isString);
                    //msExcel.SaveCurrentSheet();

                    //msExcel.SetCurrentSheet("Data");
                    //msExcel.ForceRecalc();
                    //msExcel.SaveCurrentSheet();
                    //result = true;
                }
            }

            return result;
        }

        private static int? ParseInt(string s)
        {
            int i;
            return int.TryParse(s, out i) ? (int?)i : null;
        }

        private static DateTime? ParseDate(string s)
        {
            DateTime i;
            return DateTime.TryParse(s, out i) ? (DateTime?)i : null;
        }

        private static decimal? ParseDecimal(string s)
        {
            decimal i;
            return decimal.TryParse(s, out i) ? (decimal?)i : null;
        }

        private static byte? ParseByte(string s)
        {
            byte i;
            return byte.TryParse(s, out i) ? (byte?)i : null;
        }
    }
}