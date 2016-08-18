using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace MVCWeb.Helpers
{
    public static class Helpers
    {
        public static MvcHtmlString ContentBlock1(this HtmlHelper helper)
        {
            //var sb = new StringBuilder();
            //sb.AppendFormat("<h2 class='test'>Hello World</h2>");
            return MvcHtmlString.Create( String.Format("<h2 class='test'>Hello World</h2>"));
            //sb.AppendFormat("<h2 class='{1}'>{0}</h2>", title, "Blue".Equals(GlobalProperties.color) ? "blueHeader" : string.Empty);
            //this.TextWriter.WriteLine(sb.ToString());
        }
        public static string Label1(this HtmlHelper helper, string target, string text)
        {
            return String.Format("<label for='{0}'>{1}</label>", target, text);
        }

        public static MvcHtmlString ContentBlock(this HtmlHelper helper,ISSCStyle css)
        {
            string sb = String.Format("<h2 class='{0}' style='color:{1}'>{2}</h2>",css.Class,css.color,css.Text);
           // string sb = String.Format("<p style='color:red'>This is a paragraph.</p>");
            return MvcHtmlString.Create(sb);
        }

    }
    public  class ISSCStyle
    {
        public string Class { get; set; }
        public  string color { get; set; }
        public string Text { get; set; }
    }
}