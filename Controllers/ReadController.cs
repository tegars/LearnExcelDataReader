using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelDataReader;
using LearnExcelDataReader.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace LearnExcelDataReader.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReadController : ControllerBase
    {
        [HttpGet]
        public ActionResult Get()
        {
            string text = "";
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open("Files/Book1.xlsx", FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read()) //Each ROW
                        {
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                //Console.WriteLine(reader.GetString(column));//Will blow up if the value is decimal etc. 
                                Console.WriteLine(reader.GetValue(column));//Get Value returns object
                                text+=reader.GetValue(column)+", ";
                            }
                            text +="<br/>";
                        }
                    } while (reader.NextResult()); //Move to NEXT SHEET

                }
            }
            return Ok(text);
        }
        [HttpGet("WithMapping")]
        public ActionResult WithMapping()
        {
            List<Product> products = new List<Product>();
            bool first = true;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open("Files/Book1.xlsx", FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        if (first)
                            first = false;
                        else
                        {
                            var product = new Product()
                            {
                                Name = reader.GetValue(0).ToString(),
                                Price = reader.GetValue(1).ToString(),
                                Description = reader.GetValue(2).ToString()
                            };
                            products.Add(product);
                        }
                    }
                }
            }
            return Ok(products);
        }
        [HttpPost("Upload")]
        [RequestSizeLimit(8000000)]
        public ActionResult Upload(IFormFile file)
        {
            string[] permittedExtensions = { ".xlsx",".xls" };
            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (string.IsNullOrEmpty(ext) || !permittedExtensions.Contains(ext))
                return BadRequest("Only Allowed Excel File");


            List<Product> products = new List<Product>();
            bool first = true;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var reader = ExcelReaderFactory.CreateReader(file.OpenReadStream()))
            {
                while (reader.Read())
                {
                    if (first)
                        first = false;
                    else
                    {
                        var product = new Product()
                        {
                            Name = reader.GetValue(0).ToString(),
                            Price = reader.GetValue(1).ToString(),
                            Description = reader.GetValue(2).ToString()
                        };
                        products.Add(product);
                    }
                }
            }
            return Ok(products);
        }
    }
}