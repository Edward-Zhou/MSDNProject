using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        _CommandBarButtonEvents_ClickEventHandler eventHandler;

        public static ThisAddIn instance;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //instance = this;
            PowerPoint.Application p = new PowerPoint.Application();
            
            //eventHandler = new _CommandBarButtonEvents_ClickEventHandler(App_WindowBeforeDoubleClick);

            //Globals.ThisAddIn.Application.WindowBeforeDoubleClick += new PowerPoint.EApplication_WindowBeforeDoubleClickEventHandler(App_WindowBeforeDoubleClick);// new PowerPoint.EApplication_WindowBeforeRightClickEventHandler(App_WindowBeforeDoubleClick);
            //Globals.ThisAddIn.Application.WindowBeforeRightClick += new PowerPoint.EApplication_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);// new PowerPoint.EApplication_WindowBeforeRightClickEventHandler(App_WindowBeforeDoubleClick);
            //Globals.ThisAddIn.Application.PresentationClose += Application_PresentationClose;
        }

        void Application_PresentationClose(PowerPoint.Presentation Pres)
        {
            
            //throw new NotImplementedException();
        }
        private void App_WindowBeforeDoubleClick
   (Microsoft.Office.Interop.PowerPoint.Selection Sel, ref bool Cancel)
        {
            try
            {
                this.AddItem();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error: " + exception.Message);
            }
        }
        private void App_WindowBeforeRightClick
(Microsoft.Office.Interop.PowerPoint.Selection Sel, ref bool Cancel)
        {
            try
            {
                this.AddItem1();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error: " + exception.Message);
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public void AddItem()
        {
            Microsoft.Office.Interop.PowerPoint.Application applicationObject =
      Globals.ThisAddIn.Application as Microsoft.Office.Interop.PowerPoint.Application;

            CommandBarButton commandBarButton =
       applicationObject.CommandBars.FindControl
       (MsoControlType.msoControlButton, missing, "HELLO_TAG", missing)
       as CommandBarButton;

            MessageBox.Show("WindowBeforeDoubleClick");
            commandBarButton.Click += eventHandler;
        }
        public void AddItem1()
        {
            Microsoft.Office.Interop.PowerPoint.Application applicationObject =
      Globals.ThisAddIn.Application as Microsoft.Office.Interop.PowerPoint.Application;

            CommandBarButton commandBarButton =
       applicationObject.CommandBars.FindControl
       (MsoControlType.msoControlButton, missing, "HELLO_TAG", missing)
       as CommandBarButton;

            MessageBox.Show("WindowBeforeRightClick");
            commandBarButton.Click += eventHandler;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
