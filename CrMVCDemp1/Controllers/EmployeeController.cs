using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CrMVCDemp1.Models;
using CrystalDecisions.CrystalReports.Engine;
using System.IO;
using System.Data;
using System.Reflection;

namespace CrMVCDemp1.Controllers
{
    public class EmployeeController : Controller
    {
        private EmployeeInfoEntities db = new EmployeeInfoEntities();

        // GET: Employee
        public ActionResult Index()
        {
            return View(db.Employees.ToList());
        }

        public ActionResult ExportEmployee()
        {
            ReportDocument rd = new ReportDocument();
            rd.Load(Path.Combine(Server.MapPath("~/Reports"), "CrystalReport.rpt"));
            rd.SetDataSource(ListToDatable(db.Employees.ToList()));
            Response.Buffer = false;
            Response.ClearContent();
            Response.ClearHeaders();
            try
            {
                Stream stream = rd.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                stream.Seek(0, SeekOrigin.Begin);
                return File(stream, "application/pdf", "EmloyeeList.pdf");
            }
            catch
            {
                throw;
            }
        }

        private DataTable ListToDatable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach(PropertyInfo prop  in Props)
            {
                dataTable.Columns.Add(prop.Name);
            }
            foreach(T item in items)
            {
                var values = new object[Props.Length];
                for(int i =0; i<Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
    }

}