using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace MVCWeb.Models
{
    public class Business2Context : DbContext
    {
        public DbSet<Customer> Customers { get; set; }
        public DbSet<Product> Products { get; set; }
    }

    public class Customer
    {
        public int CustomerId { get; set; }
        public string CustomerName { get; set; }
        public List<Product> Products { get; set; }

    }
    public class Product
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; }
    }
}
