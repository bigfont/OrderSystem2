using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

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
        //
        // GET: /ExcelWorkbook/

        public ActionResult Index()
        {
            var files = from file in Directory.EnumerateFiles(ExcelDirectory, "*.xlsx", SearchOption.AllDirectories)
                        select System.IO.Path.GetFileName(file);

            return View(files);
        }

        public ActionResult Worksheets(string workbook)
        {
            ViewBag.Workbook = workbook;
            var excel = new ExcelQueryFactory(Path.Combine(ExcelDirectory, workbook));
            var worksheetNames = excel.GetWorksheetNames();
            return View("Worksheets", worksheetNames);
        }

        public ActionResult Columns(string workbook, string worksheet)
        {
            ViewBag.Workbook = workbook;
            ViewBag.Worksheet = worksheet;
            var excel = new ExcelQueryFactory(Path.Combine(ExcelDirectory, workbook));
            var columnNames = excel.GetColumnNames(worksheet);
            return View(columnNames);
        }

        [HttpPost]
        public ActionResult DisplayData(IEnumerable<String> columns, string workbook, string worksheet)
        {

            var excel = new ExcelQueryFactory(Path.Combine(ExcelDirectory, workbook));
            var rows = excel.Worksheet(worksheet);

            //create data table
            DataTable dataTable = new DataTable();

            //add columns to datatable
            foreach (string c in columns)
            {
                dataTable.Columns.Add(c);
            }

            //add rows to datatable            
            foreach (LinqToExcel.Row r in rows)
            {
                DataRow row = dataTable.NewRow();
                foreach (string c in columns)
                {
                    row[c] = r[c];
                }

                bool hasData = row.ItemArray.Any(cell => 
                    cell.ToString().Length > 0 &&
                    !(cell.ToString().Equals("$ DECREASE") || cell.ToString().Equals("$ INCREASE")));

                if (hasData)
                {
                    dataTable.Rows.Add(row);
                }

            }

            return View(dataTable);
        }

    }
}
