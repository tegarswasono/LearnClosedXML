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
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public IEnumerable<WeatherForecast> Get()
        {
            var rng = new Random();
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = rng.Next(-20, 55),
                Summary = Summaries[rng.Next(Summaries.Length)]
            })
            .ToArray();
        }
        //https://www.infoworld.com/article/3538413/how-to-export-data-to-excel-in-aspnet-core-30.html
        [HttpGet("Download")]
        public ActionResult Download()
        {
            List<Author> authors = new List<Author>
            {
                new Author { Id = 1, FirstName = "Joydip", LastName = "Kanjilal" },
                new Author { Id = 2, FirstName = "Steve jos jos jos", LastName = "Smith" },
                new Author { Id = 3, FirstName = "Anand", LastName = "Narayaswamy"}
            };
            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            string fileName = "authors.xlsx";
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    GenerateBody(authors, workbook);

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, contentType, fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        private static void GenerateBody(List<Author> authors, XLWorkbook workbook)
        {
            IXLWorksheet worksheet = workbook.Worksheets.Add("Authors");
            List<string> headers = new List<string>() { "Id", "FirstName", "LastName" };
            int indexHeader = 1;
            foreach (var header in headers)
            {
                worksheet.Cell(1, indexHeader).Value = header;

                //styling
                worksheet.Range(worksheet.Cell(1, indexHeader), worksheet.Cell(3, indexHeader))
                    .Merge()
                    .Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(1, indexHeader).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                worksheet.Cell(1, indexHeader).Style.Font.SetBold();
                worksheet.Cell(1, indexHeader).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(1, indexHeader).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                indexHeader++;
            }
            
            int index = 4;
            foreach (var author in authors)
            {
                worksheet.Cell(index, 1).Value = author.Id;
                worksheet.Cell(index, 2).Value = author.FirstName;
                worksheet.Cell(index, 3).Value = author.LastName;

                //styling
                worksheet.Cell(index, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(index, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(index, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                index++;
            }

            //styling
            worksheet.Columns(1, 3).AdjustToContents();
        }
    }
    public class Author  
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }
}
