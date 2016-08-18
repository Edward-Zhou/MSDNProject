using MVCWeb.DAL;
using MVCWeb.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace MVCWeb.Controllers
{
    public class CustomerController : Controller
    {
        private Business2Context db = new Business2Context();
        private SchoolContext sc = new SchoolContext();
        public JsonResult getData()
        {
            var data = new []{
                            new
                            {
                                x = 31.949454,
                                y = 35.932913,
                                population = 50000,
                                name = "amman",
                            },
                            new
                            {
                                x = 33.79,
                                y = 33.39,
                                population = 100000,
                                name = "Zarqa",
                            }
                            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }

        // GET: Customer
        public ActionResult Index()
        {
            return View(db.Customers.ToList());
        }

        // GET: Customer/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Customer customer = db.Customers.Find(id);
            if (customer == null)
            {
                return HttpNotFound();
            }
            return View(customer);
        }

        // GET: Customer/Create
        public ActionResult Create()
        {
            ViewBag.Products = new MultiSelectList(db.Products, "ProductId", "ProductName");
            return View();
        }

        // POST: Customer/Create

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(Customer customer, int[] Products)
        {
            if (Products != null)
            {
                foreach (var ProductId in Products)
                {
                    //Product product = db.Products.Find(ProductId);
                    //customer.Products.Add(product);
                }
            }
            db.Customers.Add(customer);
            db.SaveChanges();
            return RedirectToAction("Index", db.Customers);
        }

        // GET: Customer/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Customer customer = db.Customers.Find(id);
            if (customer == null)
            {
                return HttpNotFound();
            }

            Product product = db.Products.Find(id);


            ViewBag.Products = new MultiSelectList(db.Products, "ProductId", "ProductName");

            return View(customer);
        }

        // POST: Customer/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(Customer customer, int[] Products)
        {
            if (ModelState.IsValid)
            {
                db.Entry(customer).State = EntityState.Modified;
                db.SaveChanges();
            }
            if (ModelState.IsValid)
            {
                foreach (var ProductId in Products)
                {
                    Product product = db.Products.Find(ProductId);
                    customer.Products.Add(product);
                    db.SaveChanges();
                }
            }

            return RedirectToAction("Index", db.Customers);

        }
    }
}