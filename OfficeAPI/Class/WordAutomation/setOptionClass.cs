using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeAPI.Class.WordAutomation
{
    class setOptionClass
    {
        public static void setOptions()
        {
            Word.Application wd = new Word.Application();
            wd.Selection.PasteAndFormat(WdRecoveryType.wdFormatPlainText);
            wd.Visible = true;
            wd.Documents.Add();
            wd.Options.PasteFormatBetweenDocuments = WdPasteOptions.wdMatchDestinationFormatting;
            MessageBox.Show("ok");
            
        }
        
    }
}
