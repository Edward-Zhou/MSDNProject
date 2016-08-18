using MVCWeb.DAL;
using MVCWeb.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCWeb.Controllers
{
    public class EFHomeController : Controller
    {
        private SchoolContext db = new SchoolContext();
        // GET: EFHome
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            IQueryable<EnrollmentDateGroup> data = from student in db.EFStudents
                                                   group student by student.EnrollmentDate into dategroup
                                                   select new EnrollmentDateGroup()
                                                   {
                                                       EnrollmentDate = dategroup.Key,
                                                       StudentCount = dategroup.Count()
                                                   };
            return View(data.ToList());
        }

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}