using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using PECTutorial.Models;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace PECTutorial.Controllers
{
    public class ExcelController : Controller
    {
        private IWebHostEnvironment hostingEnvironment;
        private CompanyContext context;

        public ExcelController(IWebHostEnvironment environment, CompanyContext context)
        {
            this.context = context;
            hostingEnvironment = environment;
        }

        public IActionResult ImportExcel()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ImportExcel(IFormFile excelfile)
        {
            // By old OleDbConnection way

            string path = Path.Combine(hostingEnvironment.WebRootPath, "Excel/" + excelfile.FileName);
            using (var stream = new FileStream(path, FileMode.Create))
            {
                await excelfile.CopyToAsync(stream);
            }

            string connectionString = string.Empty;
            connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;'", path);

            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                string tableName = conn.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();

                var query = $"SELECT * FROM [{tableName}]"; // The file name is used in the query
                using (var adapter = new OleDbDataAdapter(query, conn))
                {
                    var dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    List<Employee> records = dataTable.AsEnumerable().Select(row => new Employee
                    {
                        Name = row.Field<string>("Name"),// Use .Field<T>() for type safety and null handling
                        Designation = row.Field<string>("Designation"),
                        Salary = row.Field<Double>("Salary"),
                        DOB = DateTime.Parse(row.Field<string>("DOB"))
                    }).ToList();

                    context.AddRange(records);
                    context.SaveChanges();
                    ViewBag.Result = "Import Successful";
                }
            }

            return View();
        }

        public IActionResult ImportExcelOpenXml()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ImportExcelOpenXml(IFormFile excelfile)
        {
            string path = Path.Combine(hostingEnvironment.WebRootPath, "Excel/" + excelfile.FileName);
            using (var stream = new FileStream(path, FileMode.Create))
            {
                await excelfile.CopyToAsync(stream);
            }

            List<Employee> empList = new List<Employee>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                string[] header = { "Name", "Designation", "Salary", "DOB" };
                Employee emp = new Employee();
                int counter = 1;

                while (reader.Read())
                {
                    string current = reader.GetText();
                    if ((current != "") && (!header.Any(current.Contains)))
                    {
                        if (counter % 4 == 1)
                        {
                            emp.Name = current;
                        }
                        else if (counter % 4 == 2)
                        {
                            emp.Designation = current;
                        }
                        else if (counter % 4 == 3)
                        {
                            emp.Salary = double.Parse(current);
                        }
                        else
                        {
                            emp.DOB = Convert.ToDateTime(current);
                            empList.Add(emp);
                            emp = new Employee();
                        }
                        counter++;
                    }
                }
            }

            context.AddRange(empList);
            context.SaveChanges();
            ViewBag.Result = "Import Successful";

            return View();
        }

        public IActionResult ExportExcel()
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
        public IActionResult ExportExcel(List<EmployeeViewModel> empList)
        {
            var selectedRecords = empList.Where(r => r.IsChecked).Select(r => r.Id).ToList();
            var emp = context.Employee.Where(o => selectedRecords.Contains(o.Id)).ToList();

            string path = Path.Combine(hostingEnvironment.WebRootPath, "Excel/mydata.xlsx");

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {

                // Add a WorkbookPart to the document.
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Add Sheets to the Workbook.
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
                sheets.Append(sheet);

                // Add Data
                Row row = new Row();
                row.Append(new Cell() { CellValue = new CellValue("Id"), DataType = CellValues.String });
                row.Append(new Cell() { CellValue = new CellValue("Name"), DataType = CellValues.String });
                row.Append(new Cell() { CellValue = new CellValue("Destination"), DataType = CellValues.String });
                row.Append(new Cell() { CellValue = new CellValue("Salary"), DataType = CellValues.String });
                row.Append(new Cell() { CellValue = new CellValue("DOB"), DataType = CellValues.String });
                sheetData.Append(row);

                foreach (var e in emp)
                {
                    row = new Row();
                    row.Append(new Cell() { CellValue = new CellValue(e.Id), DataType = CellValues.Number });
                    row.Append(new Cell() { CellValue = new CellValue(e.Name), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(e.Designation), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(e.Salary), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(e.DOB), DataType = CellValues.Date });
                    sheetData.Append(row);
                }

                workbookPart.Workbook.Save();
            }

            var contentType = "application/octet-stream";
            return PhysicalFile(path, contentType, Path.GetFileName(path));
        }
    }
}
