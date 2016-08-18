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
using MVCWeb.ViewModels;

namespace MVCWeb.Controllers
{
    public class EFInstructorController : Controller
    {
        private SchoolContext db = new SchoolContext();

        // GET: EFInstructor
        public ActionResult Index(int? id,int? courseID)
        {
            var viewModel = new EFInstructorIndexData();
            viewModel.EFInstructors = db.EFInstructors
                .Include(i => i.EFOfficeAssignment)
                .Include(i => i.EFCourses.Select(c => c.EFDepartment))
                .OrderBy(i=>i.LastName);

            if (id != null)
            {
                ViewBag.EFInstructorID = id.Value;
                //viewModel.EFCourses = viewModel.EFInstructors.Where(i=>i.ID==id.Value).Single().EFCourses;
                if (viewModel.EFInstructors.SingleOrDefault(i => i.ID == id.Value) != null)
                {
                    viewModel.EFCourses = viewModel.EFInstructors.SingleOrDefault(i => i.ID == id.Value).EFCourses;
                }
            }

            if (courseID != null)
            {
                ViewBag.EFCourseID = courseID.Value;
                //viewModel.EFEnrollments = viewModel.EFCourses.Where(
                //    x=>x.EFCourseID==courseID.Value).Single().Enrollments;
                //lazy loading
                //if (viewModel.EFCourses.SingleOrDefault(
                //    x => x.EFCourseID == courseID.Value) != null)
                //{
                //    viewModel.EFEnrollments = viewModel.EFCourses.SingleOrDefault(
                //        x => x.EFCourseID == courseID.Value).Enrollments;
                //}

                //explicit loading

            }
            return View(viewModel);
        }

        // GET: EFInstructor/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFInstructor eFInstructor = db.EFInstructors.Find(id);
            if (eFInstructor == null)
            {
                return HttpNotFound();
            }
            return View(eFInstructor);
        }

        // GET: EFInstructor/Create
        public ActionResult Create()
        {
            ViewBag.ID = new SelectList(db.EFOfficeAssignments, "EFInstructorID", "Location");
            return View();
        }

        // POST: EFInstructor/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,LastName,FirstMidName,HireDate")] EFInstructor eFInstructor)
        {
            if (ModelState.IsValid)
            {
                db.EFInstructors.Add(eFInstructor);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.ID = new SelectList(db.EFOfficeAssignments, "EFInstructorID", "Location", eFInstructor.ID);
            return View(eFInstructor);
        }

        // GET: EFInstructor/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFInstructor eFInstructor = db.EFInstructors.Find(id);
            if (eFInstructor == null)
            {
                return HttpNotFound();
            }
            ViewBag.ID = new SelectList(db.EFOfficeAssignments, "EFInstructorID", "Location", eFInstructor.ID);
            return View(eFInstructor);
        }

        // POST: EFInstructor/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,LastName,FirstMidName,HireDate")] EFInstructor eFInstructor)
        {
            if (ModelState.IsValid)
            {
                db.Entry(eFInstructor).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.ID = new SelectList(db.EFOfficeAssignments, "EFInstructorID", "Location", eFInstructor.ID);
            return View(eFInstructor);
        }

        // GET: EFInstructor/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EFInstructor eFInstructor = db.EFInstructors.Find(id);
            if (eFInstructor == null)
            {
                return HttpNotFound();
            }
            return View(eFInstructor);
        }

        // POST: EFInstructor/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            EFInstructor eFInstructor = db.EFInstructors.Find(id);
            db.EFInstructors.Remove(eFInstructor);
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
