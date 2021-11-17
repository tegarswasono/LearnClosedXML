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
    public class ReadController : ControllerBase
    {
        private readonly ILogger<WeatherForecastController> _logger;
        public ReadController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public ActionResult<string> Get()
        {
            //var stream = System.IO.File.Open("Files/Book1.xls", FileMode.Open, FileAccess.Read);
            var workbook = new XLWorkbook("Files/Book1.xlsx");
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

            string result = "";
            foreach (var row in rows)
            {
                var rowNumber = row.RowNumber();
                var tmp = row.Cell(1).Value;
                var tmp2 = row.Cell(2).Value;
                var tmp3 = row.Cell(3).Value;

                result = result + tmp + tmp2 + tmp3 + "<br/>";
            }
            return result;
        }   
    }
}
