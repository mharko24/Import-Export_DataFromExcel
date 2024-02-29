using ExcelDataReader;
using ImportDataFromExcel.Data;
using ImportDataFromExcel.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ImportDataFromExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private string _path = "wwwroot/excel/General_DataList.xlsx";
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            List <Student> StudentList = new List<Student>();
            StudentList.Add(new Student
            {
                id = 1,
                Name = "Utenio D. Agimas",
                Address = "Camiling Tarlac"
            });
            StudentList.Add(new Student
            {
                id = 2,
                Name = "Aurelia D. Aurelio",
                Address = "San Clemente Tarlac"
            });
            return View();
        }
        public List<string> getExcelReader()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(_path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                List<string> columnData = new List<string>();

                while (reader.Read())
                {
                    if (reader.Depth >= 7) // Check if the current row is 8th row or beyond
                    {
                        var cellValue = reader.GetValue(1); // Index 1 represents column 2
                        if (cellValue != null)
                        {
                            columnData.Add(cellValue.ToString());
                        }
                    }
                }
                return columnData;
            }
        }
        public List<Student> ImportDataToExcel()
        {
            List<Student> StudentList = new List<Student>();
            StudentList.Add(new Student
            {
                id = 1,
                Name = "Utenio D. Agimas",
                Address = "Camiling Tarlac"
            });
            StudentList.Add(new Student
            {
                id = 2,
                Name = "Aurelia D. Aurelio",
                Address = "San Clemente Tarlac"
            });
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(_path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                List<string> columnData = new List<string>();

                while (reader.Read())
                {
                    if (reader.Depth >= 7) // Check if the current row is 8th row or beyond
                    {
                        var cellValue1 = reader.GetValue(3); // Index 1 represents column 2
                        var cellValue2 = reader.GetValue(4); // Index 1 represents column 2
                        var cellValue3 = reader.GetValue(5); // Index 1 represents column 2
                        var cellValue4 = reader.GetValue(6); // Index 1 represents column 2
                        foreach(var student in StudentList)
                        {
                            cellValue1 = student.id;
                            cellValue2 = student.Name;
                            cellValue3 = student.Address;

                        }
                    }
                }
                //var stream = new MemoryStream();

                
            }
            return StudentList;
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