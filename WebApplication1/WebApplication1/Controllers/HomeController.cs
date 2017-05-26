
// CORE namespaces
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

// namespace exported by NuGet package that does not support CORE
using Spire.Xls;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        // ILogger is a CORE feature
        private readonly ILogger _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            // See
            // https://www.nuget.org/packages/FreeSpire.XLS/7.9.1
            // 
            // Doco:
            // https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Worksheet/Get-a-list-of-the-worksheet-names-in-an-Excel-workbook.html

            // Here we read an .xlsx file and write its contents to the CORE logger.
            // The CORE logger in turn writes to the debug window.

            string xlsxPath = @"C:\Dev\CoreAndLegacy\test.xlsx";

            // Workbook is defined in the non-CORE package Spire.Xls
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(xlsxPath);

            Worksheet sheet = workbook.Worksheets[0];

            CellRange[] rows = sheet.Rows;
            foreach(CellRange row in rows)
            {
                _logger.LogInformation("----------- next row -----------");

                foreach (CellRange cell in row.Cells)
                {
                    string cellValue = cell.Value;
                    _logger.LogInformation(cellValue);
                }
            }

            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View();
        }
    }
}
