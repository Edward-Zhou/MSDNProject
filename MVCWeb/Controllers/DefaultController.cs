using MVCWeb.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace MVCWeb.Controllers
{
    public class DefaultController : Controller
    {
        // GET: Default
        public ActionResult Index()
        {
            var list = new List<Item>();
            list.Add(new Item { Id=1});
            list.Add(new Item { Id = 2 });
            list.Add(new Item { Id = 3 });
            list.Add(new Item { Id = 4 });

            return View(list);
        }
        public ActionResult MainView()
        {
            ViewBag.SendCodeViewModel = 1;
            ViewBag.SendCodeViewModel1 = 2;
            return View();
        }
        public ActionResult DateModel()
        {            
            return View();
        }
        [HttpPost]
        public ActionResult commanPartial()
        {
            ViewBag.Name = Request["text"].ToString();
            return View("addComment");
        }
        [HttpGet]
        public ActionResult addComment()
        {
            return View();
        }

        [HttpPost]
        public ActionResult addCommentPost()
        {
            return View();
        }
        [HttpPost]
        public ActionResult DateModel(CustomerAccount ca)
        {
            //CustomerAccount ca = new CustomerAccount();
            ca.DOB = DateTime.Now;
            return View(ca);
        }
        public ActionResult updatevat(int ID, bool status)
        {
            //Cust_Det e = (from e1 in db.Cust_Det

            //              where e1.Cust_Acc_No == ID

            //              select e1).First();

            //e.VAT_Exempt = status;

            //db.SaveChanges();
            return RedirectToAction(null, new RouteValueDictionary(new { controller = "Customer", action = "Index", store_id = -1 }));

        }

        public ActionResult updatdisc(int ID, int discuount)
        {
            //Cust_Det e = (from e1 in db.Cust_Det

            //              where e1.Cust_Acc_No == ID

            //              select e1).First();

            //e.Disc_Level = discuount;

            //db.SaveChanges();


            return RedirectToAction(null, new RouteValueDictionary(new { controller = "Customer", action = "Index", store_id = -1 }));

        }
        public ActionResult Edit(CustomerAccount model)
        {
            model=new CustomerAccount { ID = 1, Is_Biz_Cust = true, Email = "123", Full_Name = "test1" };
            //var custdetail = new MVCWeb.Models.DefaultModel.custdetail() { Company_Email="test", Company_Name="test2"};
            return View(model);
        }
        public ActionResult ExportToWord()
        {
            var list = new List<Item>();
            list.Add(new Item { Id=1});
            list.Add(new Item { Id = 2 });
            list.Add(new Item { Id = 3 });
            list.Add(new Item { Id = 4 });
            string htmlString = this.RenderRazorViewToString(@"Home\Index", list);

            Response.AddHeader("Content-Disposition", "filename=EFFReport.docx");
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