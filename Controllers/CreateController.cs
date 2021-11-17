using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace LearnClosedXML.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class CreateController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public CreateController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }
        //https://www.infoworld.com/article/3538413/how-to-export-data-to-excel-in-aspnet-core-30.html
        [HttpGet("Download")]
        public ActionResult Download()
        {
            try
            {
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string fileName = "authors.xlsx";
                var workbook = GenerateWorkbook();
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, contentType, fileName);
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        private static XLWorkbook GenerateWorkbook()
        {
            List<Author> authors = new List<Author>
                {
                    new Author { Id = 1, FirstName = "Joydip", LastName = "Kanjilal" },
                    new Author { Id = 2, FirstName = "Steve jos jos jos", LastName = "Smith" },
                    new Author { Id = 3, FirstName = "Anand", LastName = "Narayaswamy"}
                };

            var workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Authors");
            List<string> headers = new List<string>() { "Id", "FirstName", "LastName" };
            int rowHeader = 1;
            int columnHeader = 1;
            foreach (var header in headers)
            {
                worksheet.Cell(1, columnHeader).Value = header;

                //styling
                worksheet.Range(worksheet.Cell(rowHeader, columnHeader), worksheet.Cell(3, columnHeader))
                    .Merge()
                    .Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(rowHeader, columnHeader).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                worksheet.Cell(rowHeader, columnHeader).Style.Font.SetBold();
                worksheet.Cell(rowHeader, columnHeader).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(rowHeader, columnHeader).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                columnHeader++;
            }
            
            int rowBody = 4;
            foreach (var author in authors)
            {
                worksheet.Cell(rowBody, 1).Value = author.Id;
                worksheet.Cell(rowBody, 2).Value = author.FirstName;
                worksheet.Cell(rowBody, 3).Value = author.LastName;

                //styling
                worksheet.Cell(rowBody, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(rowBody, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(rowBody, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                rowBody++;
            }
            worksheet.Cell(rowBody, 1).FormulaA1 = "=SUM(A1:A" + (rowBody - 1) + ")"; //exp. =SUM(A1:A6)


            //styling
            worksheet.Columns(1, 3).AdjustToContents();
            return workbook;
        }
    }
    public class Author  
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }
}
