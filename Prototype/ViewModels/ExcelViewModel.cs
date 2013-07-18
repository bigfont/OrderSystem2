using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Prototype.ViewModels
{
    public class SimpleVendor
    {
        public int VendorID { get; set; }
        public string VendorName { get; set; }
        public List<SimpleProduct> Products { get; set; }
        public SimpleVendor()
        {
            Products = new List<SimpleProduct>();
        }
    }
    public class SimpleProduct
    {
        public string ProductName { get; set; }
        public string ProductDescription { get; set; }
    }
    public class ExcelColumnToProductPropertyMappingChoices
    {
        public List<String> ProductProperties { get; set; }
        public List<String> ExcelColumns { get; set; }

        public ExcelColumnToProductPropertyMappingChoices()
        {
            ProductProperties = new List<String>();
            ExcelColumns = new List<String>();
        }
    }
}