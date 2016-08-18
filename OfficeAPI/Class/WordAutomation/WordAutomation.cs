using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeAPI.Class.WordAutomation
{
    class WordAutomation
    {
        public static void goToBookMark()
        {
            Word.Application word = new Word.Application();
            word.Documents.Open(@"D:\OfficeDev\Word\Edward.docm",true);
            Section sction = word.ActiveDocument.Sections[1];
            sction.PageSetup.DifferentFirstPageHeaderFooter = 0;
            word.Visible = true;
            word.Selection.GoTo(WdGoToItem.wdGoToBookmark,Type.Missing,Type.Missing,"B1");
            word.Options.PasteFormatBetweenDocuments = WdPasteOptions.wdMatchDestinationFormatting;
            Style style=word.Selection.get_Style();
            word.Selection.set_Style(1);
            style.set_BaseStyle(1);
            Style basestyle = style.get_BaseStyle();
            basestyle.set_BaseStyle(1);
            
            //word.ActiveDocument.Paragraphs.First.set_Style();
            //word.ActiveDocument.Paragraphs.First.Range.sty
        }
    }
}
