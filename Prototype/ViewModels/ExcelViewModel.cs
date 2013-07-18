using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Prototype.ViewModels
{
    public class VendorProductList
    {
        public string VendorName { get; set; }
        public int VendorID { get; set; }
        public List<VendorProduct> VendorProducts { get; set; }
        public VendorProductList()
        {
            VendorProducts = new List<VendorProduct>();
        }
    }
    public class VendorProduct
    {
        public string VendorProductName { get; set; }
        public string VendorProductDescription { get; set; }
    }
    public class ExcelProductMappingChoices
    {
        public List<String> ProductProperties { get; set; }
        public List<String> ExcelColumns { get; set; }

        public ExcelProductMappingChoices()
        {
            ProductProperties = new List<String>();
            ExcelColumns = new List<String>();
        }
    }
}