using Microsoft.AspNetCore.Mvc;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Fonts;
using PECTutorial.Models;
using System.Diagnostics;

namespace PECTutorial.Controllers
{
    public class HomeController : Controller
    {
        private IWebHostEnvironment hostingEnvironment;
        private CompanyContext context;

        public HomeController(IWebHostEnvironment environment, CompanyContext context)
        {
            this.context = context;
            hostingEnvironment = environment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Pdf()
        {
            string imagePath = Path.Combine(hostingEnvironment.WebRootPath, "Images");

            var document = new Document();
            var section = document.AddSection();

            GlobalFontSettings.UseWindowsFontsUnderWindows = true;

            var heading = section.AddParagraph("Your Credit Card Statement Report has been Generated");

            heading.Format.OutlineLevel = OutlineLevel.Level1;
            heading.Format.Font.Size = 25;

            // Add line below the heading
            heading.Format.Borders.Bottom.Width = 1;
            heading.Format.SpaceAfter = "30pt";

            // Add table.
            var table = section.AddTable();

            // Add first column.
            var columnA = table.AddColumn(Unit.FromCentimeter(6));

            // Add second column.
            var columnB = table.AddColumn(Unit.FromCentimeter(12));

            // Add first row.
            var row1 = table.AddRow();

            // Add paragraph to first cell of row1.
            var cellA1 = row1[0];

            document.ImagePath = imagePath;
            var image = cellA1.AddImage("woman.jpg");
            image.Width = Unit.FromPoint(150);
            image.Height = Unit.FromPoint(150);

            // Add paragraph to second cell of row1.
            var cellB1 = row1[1];
            cellB1.AddParagraph("Name: Mrs. Grace Kelly");
            cellB1.AddParagraph("Address: House 20, 31 drowning street, London (UK)");
            cellB1.AddParagraph("Occupation: Doctor");
            cellB1.AddParagraph("Age: 30");
            cellB1.Format.Font.Size = 25;
            cellB1.Format.Font.Color = Colors.Red;

            var heading1 = section.AddParagraph("This month's transation in your Credit Card !");
            heading1.Format.Font.Size = 20;
            heading1.Format.Font.Color = Colors.BurlyWood;

            // Add line below the heading
            heading1.Format.Borders.Bottom.Width = 1;
            heading1.Format.SpaceBefore = "30pt";
            heading1.Format.SpaceAfter = "30pt";

            var table2 = section.AddTable();
            table2.Borders.Visible = true;

            table2.AddColumn("3cm");
            table2.AddColumn("3cm");
            table2.AddColumn("3cm");
            table2.AddColumn("3cm");
            table2.AddColumn("3cm");

            var row1Table2 = table2.AddRow();
            row1Table2.HeadingFormat = true;
            row1Table2.Format.Font.Color = Colors.BlueViolet;
            row1Table2.Shading.Color = Colors.LightGray;

            row1Table2[0].AddParagraph("S.No");
            row1Table2[1].AddParagraph("Merchant");
            row1Table2[2].AddParagraph("Item");
            row1Table2[3].AddParagraph("Cost");
            row1Table2[4].AddParagraph("Date");

            var row2Table2 = table2.AddRow();
            row2Table2[0].AddParagraph("1");
            row2Table2[1].AddParagraph("NYC Junction");
            row2Table2[2].AddParagraph("Fruits");
            row2Table2[3].AddParagraph("$100.00");
            row2Table2[4].AddParagraph("June 1");

            var row3Table2 = table2.AddRow();
            row3Table2[0].AddParagraph("2");
            row3Table2[1].AddParagraph("David Store");
            row3Table2[2].AddParagraph("Napkins");
            row3Table2[3].AddParagraph("5.90");
            row3Table2[4].AddParagraph("June 3");

            var row4Table2 = table2.AddRow();
            row4Table2[0].AddParagraph("3");
            row4Table2[1].AddParagraph("Singhs");
            row4Table2[2].AddParagraph("Toys");
            row4Table2[3].AddParagraph("$99.99");
            row4Table2[4].AddParagraph("June 9");

            var row5Table2 = table2.AddRow();
            row5Table2[0].AddParagraph("4");
            row5Table2[1].AddParagraph("Seven 11");
            row5Table2[2].AddParagraph("Grocery");
            row5Table2[3].AddParagraph("$140.00");
            row5Table2[4].AddParagraph("June 15");

            var row6Table2 = table2.AddRow();
            row6Table2[0].AddParagraph("5");
            row6Table2[1].AddParagraph("Carlos Pharmacy");
            row6Table2[2].AddParagraph("Drugs");
            row6Table2[3].AddParagraph("$60.00");
            row6Table2[4].AddParagraph("June 25");

            var custName = section.AddParagraph("Hello Grace,");
            custName.Format.SpaceBefore = "30pt";
            custName.Format.SpaceAfter = "20pt";
            section.AddParagraph("Thank you for being our valuable customer. We hope our letter finds you in the best of health and wealth.\n\nYours Sincerely.\nICICI Bank");

            // Create a PDF renderer for the MigraDoc document.
            var pdfRenderer = new PdfDocumentRenderer();

            // Associate the MigraDoc document with a renderer.
            pdfRenderer.Document = document;

            // Layout and render document to PDF.
            pdfRenderer.RenderDocument();
            // Save the document.
            pdfRenderer.Save("SimpleDocument.pdf");

            return RedirectToAction("Index");
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
