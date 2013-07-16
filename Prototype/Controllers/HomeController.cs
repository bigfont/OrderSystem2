using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Prototype.Models;

namespace Prototype.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Message = "Modify this template to jump-start your ASP.NET MVC application.";

            using (var db = new OrderSystemContext())
            { 
                //creat objects
                Vendor vendor = new Vendor();
                vendor.VendorName = "Test Vendor";

                Product product = new Product();
                product.Vendor = vendor;
                product.ProductDescription = "Test product";

                //add objects to context
                db.Vendors.Add(vendor);
                db.Products.Add(product);

                //save context
                db.SaveChanges();

                //make a query
                var query = from v in db.Vendors
                            orderby v.VendorName
                            select v;

                //execute the query
                foreach (var v in query)
                {
                    var name = v.VendorName;
                }

                //done
            }

            return View();
        }
    }
}
