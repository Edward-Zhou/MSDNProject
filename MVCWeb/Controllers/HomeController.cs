using MVCWeb.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCWeb.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult ExportToWord()
        {
            var list = new List<Item>();
            list.Add(new Item { Id = 1 });
            list.Add(new Item { Id = 2 });
            list.Add(new Item { Id = 3 });
            list.Add(new Item { Id = 4 });
            string htmlString = this.RenderRazorViewToString(@"Index", list);
            Response.AddHeader("Content-Disposition", "filename=EFFReport.doc");
            Response.ContentType = "application/msword";
            Response.Write(htmlString);
            return null;
        }

        public string RenderRazorViewToString(string viewName, object model)
        {
            ViewData.Model = model;
            using (var sw = new StringWriter())
            {
                var viewResult = ViewEngines.Engines.FindPartialView(ControllerContext,
                                                                         viewName);
                var viewContext = new ViewContext(ControllerContext, viewResult.View,
                                             ViewData, TempData, sw);
                viewResult.View.Render(viewContext, sw);
                viewResult.ViewEngine.ReleaseView(ControllerContext, viewResult.View);
                return sw.GetStringBuilder().ToString();
            }
        }


    }
}