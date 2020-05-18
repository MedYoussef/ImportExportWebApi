using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ExcelHandler.Models;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace ExcelHandler.Helpers
{

    public static class ExcelHelper
    {
        public static async Task<ExcelResponse<List<User>>> ReadFile(IFormFile formFile, CancellationToken cancellationToken)
        {
            var list = new List<User>();
            var errors = new List<Error>();
            var sex = new List<string> {"Male", "Female", "Other" };
            using (var stream = new MemoryStream())
            {
                await formFile.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;
                    int a;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        if(worksheet.Cells[row, 1].Value.ToString().Trim().Length > 20)
                        {
                            errors.Add(new Error { Row = row, Message = "The Name is too long" });
                        }
                        if(!sex.Contains(worksheet.Cells[row, 2].Value.ToString().Trim()))
                        {
                            errors.Add(new Error { Row = row, Message = "The Sex must be either Male, Female or Other" });
                        }
                        if(!int.TryParse(worksheet.Cells[row, 3].Value.ToString().Trim(), out a) || a < 0)
                        {
                            errors.Add(new Error { Row = row, Message = "The Age must be an integer greater then 0" });
                        }
                        if(errors.Count == 0)
                        {
                            list.Add(new User
                            {
                                Name = worksheet.Cells[row, 1].Value.ToString().Trim(),
                                Sex = worksheet.Cells[row, 2].Value.ToString().Trim(),
                                Age = int.Parse(worksheet.Cells[row, 3].Value.ToString().Trim()),
                            });
                        }
                        
                    }
                }
            }
            return errors.Count > 0 ?
                ExcelResponse<List<User>>.GetResult(-2, errors, list) :
                ExcelResponse<List<User>>.GetResult(0, errors, list);
        }
    }
}
