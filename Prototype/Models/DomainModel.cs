using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace Prototype.Models
{
    public class Vendor
    {
        public int VendorID { get; set; }
        public string VendorName { get; set; }
        public virtual List<Product> Products { get; set; }
    }
    public class Product
    {
        public int ProductID { get; set; }
        public string ProductDescription { get; set; }
        public int VendorID { get; set; }
        public virtual Vendor Vendor { get; set; }
    }
    public class OrderSystemContext : DbContext
    {
        public DbSet<Vendor> Vendors { get; set; }
        public DbSet<Product> Products { get; set; }
    }
}