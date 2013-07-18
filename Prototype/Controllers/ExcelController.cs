using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Prototype.Models;
using Prototype.ViewModels;
using System.Reflection;

namespace Prototype.Controllers
{
    public class ExcelController : Controller
    {
        private string ExcelDirectory
        {
            get
            {
                return Server.MapPath("~/ExcelWorkbooks");
            }
        }

        public ActionResult Index()
        {
            IEnumerable<System.String> workbookNames = from file in Directory.EnumerateFiles(ExcelDirectory, "*.xlsx", SearchOption.AllDirectories)
                                                       select System.IO.Path.GetFileName(file);
            return View(workbookNames);
        }

        public ActionResult Worksheets(string workbookName)
        {
            ViewBag.WorkbookName = workbookName;
            var excel = new ExcelQueryFactory(Path.Combine(ExcelDirectory, workbookName));
            IEnumerable<System.String> worksheetNames = excel.GetWorksheetNames();
            return View("Worksheets", worksheetNames);
        }

        public ActionResult Columns(string workbookName, string worksheetName)
        {
            ViewBag.WorkbookName = workbookName;
            ViewBag.WorksheetName = worksheetName;
            var excel = new ExcelQueryFactory(Path.Combine(ExcelDirectory, workbookName));
            IEnumerable<System.String> excelColumnNames = excel.GetColumnNames(worksheetName);

            ExcelProductMappingChoices mappingChoices = new ExcelProductMappingChoices();
            mappingChoices.ExcelColumns = excelColumnNames.ToList<String>();
            PropertyInfo[] vendorProducts = typeof(VendorProduct).GetProperties();
            foreach (PropertyInfo pi in vendorProducts)
            {
                mappingChoices.ProductProperties.Add(pi.Name);
            }

            return View(mappingChoices);
        }

        [HttpPost]
        public ActionResult DisplayData(string workbookName, string worksheetName, VendorProduct mappings)
        {
            string[] stringsToAvoid = { "$ DECREASE", "$ INCREASE", "NEW", "DISCONTINUED" };

            ViewBag.WorkbookName = workbookName;
            ViewBag.WorksheetName = worksheetName;

            //TODO Can we use a typed form or pass a typed object?
            var excel = new ExcelQueryFactory(Path.Combine(ExcelDirectory, workbookName));
            var rows = excel.Worksheet(worksheetName);

            ////create vendor product list
            VendorProductList vendorProductList = new VendorProductList();
            vendorProductList.VendorName = "test vendor";
            vendorProductList.VendorID = 0;

            //add rows to datatable            
            foreach (LinqToExcel.Row r in rows)
            {
                VendorProduct product = new VendorProduct();

                product.VendorProductName = r[mappings.VendorProductName];
                product.VendorProductDescription = r[mappings.VendorProductDescription];

                PropertyInfo[] properties = typeof(VendorProduct).GetProperties();
                bool hasData = properties.Any<PropertyInfo>(pi =>
                    pi.GetValue(product) != null &&
                    pi.GetValue(product).ToString().Length > 0 &&
                    !stringsToAvoid.Any<String>(pi.GetValue(product).ToString().Contains));                

                if (hasData)
                {
                    vendorProductList.VendorProducts.Add(product);
                }
            }

            return View(vendorProductList);
        }

        [HttpPost]
        public ActionResult ImportData(IEnumerable<VendorProduct> products)
        {       
            //TODO 
            //Capture the POST from the from into an object
            //Persist the data to the database
            //The problem is doing a strongly-typed capture of the form data.

            return View();
        }
    }
}
