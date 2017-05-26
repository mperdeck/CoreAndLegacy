using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Spire.Xls;
using System.Windows.Forms;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            // See
            // https://www.nuget.org/packages/FreeSpire.XLS/7.9.1
            // 
            // Doco:
            // https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Worksheet/Get-a-list-of-the-worksheet-names-in-an-Excel-workbook.html

            string xlsxPath = @"C:\Dev\CoreAndLegacy\test.xlsx";

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(xlsxPath);

            Worksheet sheet = workbook.Worksheets[0];
            var dataGridView = new DataGridView();

            dataGridView.DataSource = sheet.ExportDataTable(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn, false);

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    string value = cell.Value.ToString();

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
