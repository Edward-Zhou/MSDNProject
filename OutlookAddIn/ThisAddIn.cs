using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        internal Helper helper;
        internal Outlook.Items items;
        internal Outlook.MailItem item;
        internal Outlook.AppointmentItem appointmentItem;
        string _PreviousMeetingId = string.Empty; 
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            helper = new Helper();
            items = Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).Items;
            items.ItemAdd += items_ItemAdd;
        }
        
        void items_ItemAdd(object Item)
        {
            appointmentItem = items.GetLast() as Outlook.AppointmentItem;
            //string sub = appointmentItem.Subject;
            //_PreviousMeetingId = appointmentItem.GlobalAppointmentID;
            var status = appointmentItem.MeetingStatus;
            //MessageBox.Show(appointmentItem.Subject);

            //IQueryable<Outlook.AppointmentItem> AItems=items.AsQueryable();
            //appointmentItem = Item as Outlook.AppointmentItem;
            //Outlook.AppointmentItem item = AItems.Where(a=>a.Subject==appointmentItem.Subject);
            //try
            //{                
            //    appointmentItem = Item as Outlook.AppointmentItem;
            //    _PreviousMeetingId = appointmentItem.GlobalAppointmentID;
            //    MessageBox.Show(_PreviousMeetingId);

            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show(e.InnerException.ToString());
            //}
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }
        public void CreateHtmlFolder()
        {
            Outlook.MAPIFolder newView = null;
            string viewName = "HtmlView";
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)
                this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook
                .OlDefaultFolders.olFolderInbox);
            Outlook.Folders searchFolders = (Outlook.Folders)inBox.Folders;
            bool foundView = false;
            foreach (Outlook.MAPIFolder searchFolder in searchFolders)
            {
                if (searchFolder.Name == viewName)
                {
                    newView = inBox.Folders[viewName];
                    foundView = true;
                }
            }
            if (!foundView)
            {
                newView = (Outlook.MAPIFolder)inBox.Folders.
                    Add("HtmlView", Outlook.OlDefaultFolders.olFolderInbox);
                newView.WebViewURL = "https://microsoft.sharepoint.com/teams/msdnofficedev/Shared%20Documents/Forms/AllItems.aspx";
                newView.WebViewOn = true;
            }
            Application.ActiveExplorer().SelectFolder(newView);
            Application.ActiveExplorer().CurrentFolder.Display();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            ((Outlook.ApplicationEvents_11_Event)(this.Application)).Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Quit);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }


        private void ThisAddIn_Quit()
        {
            helper.CloseConnection();
        }
        
        #endregion
    }
}
