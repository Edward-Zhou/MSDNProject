using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MVCWeb.DAL;
using MVCWeb.Models;
using PagedList;

namespace MVCWeb.Controllers
{
    public class EFStudentController : Controller
    {
        private SchoolContext db = new SchoolContext();

        // GET: EFStudent
        public ActionResult Index(string sortOrder,string currentFilter,string searchString,int?page)
        {
            ViewBag.CurrentSort = sortOrder;
            ViewBag.NameSortParm = String.IsNullOrEmpty(sortOrder) ? "name_desc" : "";
            ViewBag.DateSortParm = sortOrder == "Date" ? "date_desc" : "Date";
            if (searchString != null)
            {
                page = 1;
            }
            else
            {
                searchString = currentFilter;
            }
            ViewBag.CurrentFilter = searchString;
            var students = from s in db.EFStudents
                           select s;
            if (!String.IsNullOrEmpty(searchString))
            {
                students = students.Where(s=>s.LastName.Contains(searchString)||
                    s.FirstMidName.Contains(searchString));
            }
            switch (sortOrder)
            { 
                case "name_desc":
                    students = students.OrderByDescending(s=>s.LastName);
                    break;
                case "Date":
                    students = students.OrderBy(s=>s.EnrollmentDate);
                    break;
                case "date_desc":
                    students = students.OrderByDescending(s=>s.EnrollmentDate);
                    break;
                default:
                    students = students.OrderBy(s=>s.LastName);
                    break;
            }
            int pageSize = 3;
            int pageNumber=(page??1);

            return View(students.ToPagedList(pageNumber,pageSize));
        }

        // GET: EFStudent/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFStudent eFStudent = db.EFStudents.Find(id);
            if (eFStudent == null)
            {
                return HttpNotFound();
            }
            return View(eFStudent);
        }

        // GET: EFStudent/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: EFStudent/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        //ValidateAntiForgeryToken attribute helps prevent cross-site request forgery. it requires a corresponding Html.AntiForgeryToken()
        //statement in the view
        [ValidateAntiForgeryToken]
        //remove this id, because id is primary, and it created by sql
        //Bind is used for depending which fields to create
        public ActionResult Create([Bind(Include = "LastName,FirstMidName,EnrollmentDate")] EFStudent eFStudent)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    db.EFStudents.Add(eFStudent);
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
            }
            catch (DataException)
            {
                //Log the error (uncomment dex variable name and add a line here to write a log.
                ModelState.AddModelError("", "Unable to save changes. Try again, and if the problem persists see your system administrator.");
            }

            return View(eFStudent);
        }

        // GET: EFStudent/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFStudent eFStudent = db.EFStudents.Find(id);
            if (eFStudent == null)
            {
                return HttpNotFound();
            }
            return View(eFStudent);
        }

        // POST: EFStudent/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,LastName,FirstMidName,EnrollmentDate")] EFStudent eFStudent)
        {
            if (ModelState.IsValid)
            {
                db.Entry(eFStudent).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(eFStudent);
        }

        // GET: EFStudent/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFStudent eFStudent = db.EFStudents.Find(id);
            if (eFStudent == null)
            {
                return HttpNotFound();
            }
            return View(eFStudent);
        }

        // POST: EFStudent/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            EFStudent eFStudent = db.EFStudents.Find(id);
            db.EFStudents.Remove(eFStudent);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
