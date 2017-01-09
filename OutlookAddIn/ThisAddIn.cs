using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;
using Microsoft.Office.Core;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        internal Helper helper;
        internal Outlook.Items items;
        internal Outlook.MailItem item;
        internal Outlook.AppointmentItem appointmentItem;
        internal Outlook.TaskItem taskItem;
        string _PreviousMeetingId = string.Empty; 
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.NewMailEx += Application_NewMailEx;
            //Application.ItemContextMenuDisplay += Application_ItemContextMenuDisplay;
            //helper = new Helper();
            //items = Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).Items;
            //items.ItemAdd += items_ItemAdd;
        }

        void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            var contextButton = CommandBar.Controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true) as CommandBarButton;
            contextButton.Visible = true;
            contextButton.Caption = "some caption...";
            contextButton.Click += new _CommandBarButtonEvents_ClickEventHandler(contextButton_Click);

        }
        void contextButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("ok");
        }
        void Application_NewMailEx(string EntryIDCollection)
        {
            string TorgateMail = "v-tazho@hotmail.com";
            string[] entryIDs = EntryIDCollection.Split(',');
            foreach (string entryID in entryIDs)
            {
                Outlook.NameSpace ns = Application.Session;
                Outlook.MailItem newmail = ns.GetItemFromID(entryID, missing) as Outlook.MailItem;
                if (newmail.SenderEmailAddress.ToLower() == TorgateMail)
                {
                    newmail.Delete();
                }
                //int b1 = String.Compare(TorgateMail, newmail.SenderEmailAddress, true);
                //if (b1 == 0)
                //{
                //    newmail.Delete();
                //}
            }
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

        //protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        //{
        //    return new RibbonContextMenu();
        //}
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
