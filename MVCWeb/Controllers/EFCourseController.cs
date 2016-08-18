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

namespace MVCWeb.Controllers
{
    public class EFCourseController : Controller
    {
        private SchoolContext db = new SchoolContext();

        // GET: EFCourse
        public ActionResult Index()
        {
            var eFCourses = db.EFCourses.Include(e => e.EFDepartment);
            return View(eFCourses.ToList());
        }

        // GET: EFCourse/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFCourse eFCourse = db.EFCourses.Find(id);
            if (eFCourse == null)
            {
                return HttpNotFound();
            }
            return View(eFCourse);
        }

        // GET: EFCourse/Create
        public ActionResult Create()
        {
            ViewBag.EFDepartmentID = new SelectList(db.EFDepartments, "EFDepartmentID", "Name");
            return View();
        }

        // POST: EFCourse/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "EFCourseID,Title,Credits,EFDepartmentID")] EFCourse eFCourse)
        {
            if (ModelState.IsValid)
            {
                db.EFCourses.Add(eFCourse);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.EFDepartmentID = new SelectList(db.EFDepartments, "EFDepartmentID", "Name", eFCourse.EFDepartmentID);
            return View(eFCourse);
        }

        // GET: EFCourse/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFCourse eFCourse = db.EFCourses.Find(id);
            if (eFCourse == null)
            {
                return HttpNotFound();
            }
            ViewBag.EFDepartmentID = new SelectList(db.EFDepartments, "EFDepartmentID", "Name", eFCourse.EFDepartmentID);
            return View(eFCourse);
        }

        // POST: EFCourse/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "EFCourseID,Title,Credits,EFDepartmentID")] EFCourse eFCourse)
        {
            if (ModelState.IsValid)
            {
                db.Entry(eFCourse).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.EFDepartmentID = new SelectList(db.EFDepartments, "EFDepartmentID", "Name", eFCourse.EFDepartmentID);
            return View(eFCourse);
        }

        // GET: EFCourse/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFCourse eFCourse = db.EFCourses.Find(id);
            if (eFCourse == null)
            {
                return HttpNotFound();
            }
            return View(eFCourse);
        }

        // POST: EFCourse/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            EFCourse eFCourse = db.EFCourses.Find(id);
            db.EFCourses.Remove(eFCourse);
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
