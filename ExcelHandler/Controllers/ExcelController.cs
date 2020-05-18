using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelHandler.Helpers;
using ExcelHandler.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ExcelHandler.Controllers
{
    [Route("api/[controller]")]
    public class ExcelController : Controller
    {
        [HttpPost("import")]
        public async Task<ExcelResponse<List<User>>> Import(IFormFile formFile, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                //return ExcelResponse<List<User>>.GetError("The file is empty");
                return ExcelResponse<List<User>>.GetError("The File is Empty");
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelResponse<List<User>>.GetError("Not Supported file extension");
            }

            //List<User> usersList = await ExcelHelper.ReadFile(formFile, cancellationToken);

            // We Can Add The list of users to Db
            return await ExcelHelper.ReadFile(formFile, cancellationToken);
            //return ExcelResponse<List<User>>.GetResult(0, "OK", usersList);
        }

        [HttpGet("export")]
        public async Task<IActionResult> Export(CancellationToken cancellationToken)
        {
            await Task.Yield();
            var list = new List<User>
            {
               new User { Name = "Youssef", Sex="Male", Age = 26},
               new User { Name = "Mohamed", Sex = "Male", Age = 25}
            };
            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells["A1:C1"].Style.Font.Bold = true;
                workSheet.Cells.LoadFromCollection(list, true);
                package.Save();
            }
            stream.Position = 0;
            string excelName = $"UserList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }

    }
}
