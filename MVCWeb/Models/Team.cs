using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCWeb.Models
{
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
    public class Items
    {
        public List<Item> item { get; set; }
    }

}