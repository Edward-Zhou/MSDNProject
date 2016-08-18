using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAPI.Class.OpenXmlHelper.ExcelHelper
{
    class CellData
    {
        private void refreshDictionaryOpenXML()
        {
            //lblRefreshMessage.Text = "Refreshing Dictionary,Please Wait...";
            //btnTranslate.Enabled = false;

            string strDoc = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\dict1.xlsx";

            using (SpreadsheetDocument spreadsheetDocument =
                        SpreadsheetDocument.Open(strDoc, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                int str;
                int keyVal;

                try
                {

                    foreach (Row r in sheetData.Elements<Row>())
                    {
                        int colNo = 1;
                        string strString = "";
                        string keyString = "";
                        foreach (Cell c in r.Elements<Cell>())
                        {
                            if (colNo >= 3)
                            {
                                break;
                            }
                            else if (colNo == 1)
                            {
                                if (c.CellValue != null && Int32.TryParse(c.InnerText, out str))
                                {
                                }
                                else
                                {
                                    str = -1;

                                }

                                if (str != -1)
                                {

                                    //strString = GetSharedStringItemById(workbookPart, str);
                                    //byte[] bytes = Encoding.Default.GetBytes(strString);
                                    //byte[] devBytes = Encoding.Convert(Encoding.Default, encoding, bytes);
                                    //strString = encoding.GetString(devBytes);

                                }

                            }
                            else if (colNo == 2)
                            {
                                if (c.CellValue != null && Int32.TryParse(c.InnerText, out keyVal))
                                {
                                }
                                else
                                {
                                    keyVal = -1;

                                }

                                if (keyVal != -1)
                                {
                                    //keyString = GetSharedStringItemById(workbookPart, keyVal);

                                }
                            }
                            colNo += 1;
                        }

                        //if (!translatorDictionary.ContainsKey(strString))
                        //{
                        //    translatorDictionary.Add(strString, keyString);

                        //}

                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }
            //lblRefreshMessage.Text = "Dictionary Refreshed";
            //btnTranslate.Enabled = true;
        }
    }
}
