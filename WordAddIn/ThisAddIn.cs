using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.IO;
using log4net.Repository.Hierarchy;
using log4net.Appender;
using System.Windows.Forms;
using System.Runtime.InteropServices;
[assembly: log4net.Config.XmlConfigurator(ConfigFile = "App.config", Watch = true)]   
namespace WordAddIn
{
    public partial class ThisAddIn
    {
        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public static Word.Application WordApp;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            this.Application.DocumentOpen += Application_DocumentOpen;
            //this.Application.DocumentBeforeClose+=Application_DocumentBeforeClose;
            
            //WordApp = this.Application;
            ////MessageBox.Show(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            //string Name = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\TestSaveAs.docm";

            //WordApp.ActiveDocument.SaveAs(Name,
            //        Type.Missing, Type.Missing, Type.Missing,
            //        Type.Missing, Type.Missing, Type.Missing,
            //        Type.Missing, Type.Missing, Type.Missing,
            //        Type.Missing, Type.Missing, Type.Missing,
            //        Type.Missing, Type.Missing, Type.Missing);
            //MessageBox.Show("Add in load");
            string filePath = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
            log4net.Config.XmlConfigurator.Configure(new FileInfo(filePath));

            //log4net.Config.XmlConfigurator.Configure(new FileInfo(@"D:\Edward\Project\MSDNProject\MSDNProject\WordAddIn\App.config"));
            // Get the Hierarchy object that organizes the loggers
            Hierarchy hier = log4net.LogManager.GetRepository() as Hierarchy;

            if (hier != null)
            {
                // Get ADONetAppender
                var rollingFileAppender =
                    (RollingFileAppender)hier.GetAppenders().Where(
                        appender => appender.Name.Equals("RollingLogFileAppender", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();

                if (rollingFileAppender != null)
                {
                    string file = Path.GetDirectoryName(filePath) +@"\log.txt";
                    if (!File.Exists(file))
                    {
                        File.Create(file);
                    }
                    rollingFileAppender.File = file;
                    //rollingFileAppender.File = @"D:\log\log.txt";
                    rollingFileAppender.ActivateOptions();                   
                }
            }
            myUserControl1 = new MyUserControl();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            //if you do not show up "my task pane", you could use the code below
            //myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, " ");
            myCustomTaskPane.Visible = true;
            myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            //Form1 f1 = new Form1();
            //f1.Show();
            //change height 
            //myCustomTaskPane.Height = 60;
            //this.Application.ActiveDocument.SaveAs2(FileName: "your file path", FileFormat: WdSaveFormat.wdFormatDocument);
            //this.Application.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(App_WindowSelectionChangeEventHandler);
            
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {

            //object activeWindow = null;
            //try
            //{
            //    activeWindow = Doc.ActiveWindow;// Doc.GetType().InvokeMember("ActiveWindow", System.Reflection.BindingFlags.GetProperty, null, Doc, null);
                
            //}
            //catch (Exception)
            //{
            //    // skip the exception
            //}
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(activeWindow);
            //Doc = null;
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(Doc);

            //string message = "(not set)";
            //if (activeWindow != null)
            //{
            //    message = "Caption = " + (activeWindow as Word.Window).Caption;
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeWindow);
            //}
            //else
            //{
            //    message = "no active window";
            //}
            System.Diagnostics.Debug.WriteLine("!!! " + Doc.ActiveWindow.Caption);
            //this.Application.Documents[Doc.ActiveWindow.Caption].Close();
        }

        private void App_WindowSelectionChangeEventHandler(Selection Sel)
        {
            MessageBox.Show(Sel.Text);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //Docu = null;
            //Docu.ActiveWindow.Close();
            this.Application.DocumentOpen -= Application_DocumentOpen;
            //Marshal.ReleaseComObject(this.Application);
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon5();
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
