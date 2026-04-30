using Microsoft.AspNetCore.Mvc;
using PECTutorial.Models;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace PECTutorial.Controllers
{
    public class CsvController : Controller
    {
        private IWebHostEnvironment hostingEnvironment;
        private CompanyContext context;

        public CsvController(IWebHostEnvironment environment, CompanyContext context)
        {
            this.context = context;
            hostingEnvironment = environment;
        }

        public IActionResult ImportCsv()
        {
            return View();
        }


        [HttpPost]
        public async Task<IActionResult> ImportCsv(IFormFile csvfile)
        {
            // By old OleDbConnection way

            string path = Path.Combine(hostingEnvironment.WebRootPath, "CSV/" + csvfile.FileName);
            using (var stream = new FileStream(path, FileMode.Create))
            {
                await csvfile.CopyToAsync(stream);
            }

            string folderPath = Path.Combine(hostingEnvironment.WebRootPath, "CSV"); // The Data Source is the folder, not the file
            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={folderPath};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1;MaxScanRows=0""";

            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                var query = $"SELECT * FROM [{csvfile.FileName}]"; // The file name is used in the query
                using (var adapter = new OleDbDataAdapter(query, conn))
                {
                    var dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    List<Employee> records = dataTable.AsEnumerable().Select(row => new Employee
                    {
                        Name = row.Field<string>("Name"),// Use .Field<T>() for type safety and null handling
                        Designation = row.Field<string>("Designation"),
                        Salary = row.Field<Double>("Salary"),
                        DOB = row.Field<DateTime>("DOB")
                    }).ToList();

                    context.AddRange(records);
                    context.SaveChanges();
                    ViewBag.Result = "Import Successful";
                }
            }

            return View();
        }

        public IActionResult ExportCsv()
        {
            List<Employee> eList = context.Employee.ToList();

            List<EmployeeViewModel> records = eList.AsEnumerable().Select(row => new EmployeeViewModel
            {
                Id = row.Id,
                Name = row.Name,
                Designation = row.Designation,
                Salary = row.Salary,
                DOB = row.DOB
            }).ToList();

            return View(records);
        }

        [HttpPost]
        public IActionResult ExportCsv(List<EmployeeViewModel> empList)
        {
            var selectedRecords = empList.Where(r => r.IsChecked).Select(r => r.Id).ToList();
            var emp = context.Employee.Where(o => selectedRecords.Contains(o.Id)).ToList();

            var sb = new StringBuilder();

            sb.Append("Id" + ',' + "Name" + ',' + "Designation" + ',' + "Salary" + ',' + "DOB"); // header
            sb.Append("\r\n"); // New line after header

            foreach (var e in emp)
            {
                sb.Append(e.Id.ToString() + ',' + e.Name + ',' + e.Designation + ',' + e.Salary + ',' + e.DOB);
                sb.Append("\r\n");
            }

            return File(Encoding.UTF8.GetBytes(sb.ToString()), "text/csv", "exportdata.csv");
        }
    }
}
