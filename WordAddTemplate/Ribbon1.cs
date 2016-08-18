using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
namespace WordAddTemplate
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        object LockComments = false;
        object AddToRecentFiles = true;
        object ReadOnlyRecommended = false;
        object EmbedTrueTypeFonts = false;
        object SaveNativePictureFormat = true;
        object SaveFormsData = true;
        object SaveAsAOCELetter = false;
        object Encoding = MsoEncoding.msoEncodingUSASCII;
        object InsertLineBreaks = false;
        object AllowSubstitutions = false;
        object LineEnding = Word.WdLineEndingType.wdCRLF;
        object AddBiDiMarks = false;
        object wdCompatibilityMode = 15;

        private void SaveAsTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisDocument.bookmark1.Text = "test";
            object FileName = System.IO.Path.GetDirectoryName(
      System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + @"\myfile.dotx";
            object FileFormat = Word.WdSaveFormat.wdFormatXMLTemplate;
            object missing = Type.Missing;            
            Globals.ThisDocument.SaveAs(ref FileName, ref FileFormat, ref missing,
            ref missing, ref missing, ref missing,
            ref missing, ref missing,
            ref missing, ref missing,
            ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);
            MessageBox.Show("ok");
        }
    }
}
