using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word=Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutlookAddIn
{
    public partial class Ribbon2
    {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void MailStatus_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application _application=Globals.ThisAddIn.Application;
            
            string rootFolderPath = _application.Session.DefaultStore.GetRootFolder().FolderPath;
            Outlook.NameSpace outlookNameSpace = _application.GetNamespace("MAPI");
            Outlook.MAPIFolder sentfolder = _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail) as Outlook.MAPIFolder;
               string filter = "@SQL=" + "\""
                                + "urn:schemas:httpmail:textdescription" + "\""
                                + " like '%{0}%'";
               string query = string.Format(filter, "c34912a4-8318-464b-b50f-a0cda81f44bf");
           //var mailItem = sentfolder.Items.Find(query);
           var mailItem = _application.Application.ActiveExplorer().Selection[1];
           if (mailItem is Outlook.MailItem)
            {
                                var mail = (Outlook.MailItem)mailItem;
                                var status1 =mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10800003");

             var status2 = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/id/{0006200B-0000-0000-C000-000000000046}/88090003");
            }


        }

        private void WordRange_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application wapp=new Word.Application();
            wapp.Visible = true;
            Word.Document doc= wapp.Documents.Open(@"D:\OfficeDev\Word\201602\ChineseString.docm");
            Word.Range r=doc.Range(wapp.Selection.Start,wapp.Selection.End);
            r.Text="test";

        }

        private void ItemAddbtn_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Items items =Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).Items;
            Outlook.AppointmentItem ai = items.GetLast() as Outlook.AppointmentItem;
            string sub = ai.Subject;
        }

        private void ShowForm_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (Microsoft.Office.Tools.Outlook.IFormRegion formRegion
        in Globals.FormRegions)
            {
                if (formRegion is FormRegion2)
                {
                    FormRegion2 formRegion1 = (FormRegion2)formRegion;
                    formRegion1.Visible = !formRegion1.Visible;
                    //formRegion1.Location
                    //formRegion1.textBox1.Text = "Hello World";
                }
            }

            //WindowFormRegionCollection formRegion=Globals.FormRegions[Globals.ThisAddIn.Application.ActiveInspector()];
            //formRegion.FormRegion2.Visible = !formRegion.FormRegion2.Visible;
        
        }

        private void ShowWebView_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CreateHtmlFolder();
        }
        int propKey=0;
        private void SetPropertybtn_Click(object sender, RibbonControlEventArgs e)
        {
          
                MailItem mail =(MailItem)Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
                if (mail != null)
                {
                    mail.UserProperties.Add(propKey.ToString(), OlUserPropertyType.olText, true, OlFormatText.olFormatTextText);
                    mail.UserProperties[propKey.ToString()].Value = propKey.ToString();
                    mail.Save();
                    Marshal.ReleaseComObject(mail);
                }
                propKey += 1;
            }

        }


    }

