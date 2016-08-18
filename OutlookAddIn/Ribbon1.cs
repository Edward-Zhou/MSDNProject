using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OutlookAddIn
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        public void CallMethod()
        {
            
            MessageBox.Show("ok");
        }
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
        private bool enable=false;
        private Office.IRibbonUI rUI;
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public void OnAction(Office.IRibbonControl control)
        {
            Console.WriteLine("ok");
            //this.rUI;

        }
        public void OnAction1(Office.IRibbonControl control)
        {
            if (enable == true)
            {
                enable = false;
            }
            else
            {
                enable = true;
            }

            Console.WriteLine("ok1");
        }
        public bool GetEnable(Office.IRibbonControl control)
        {
            return enable;
          
        }
        int propKey = 0;
        public void SetCustomProperty(Office.IRibbonControl control)
        {

            if (control.Context is Selection)
            {
                Selection sel = control.Context as Selection;
                MailItem mail = (MailItem)sel[1];
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
        #endregion
        
        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
