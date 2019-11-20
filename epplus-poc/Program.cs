using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace EpPlusPoC
{
    /// <summary>
    /// based on https://stackoverflow.com/a/37746915/725987
    /// </summary>
    internal static class Program
    {
        private static void Main()
        {
            using (var fileStream = new FileStream("sample-data/financial-sample.xlsx", FileMode.Open))
            {
                var excelPackage = new ExcelPackage(fileStream);
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                
                IEnumerable<ExcelResourceDto> objectsList = worksheet.ConvertSheetToObjects<ExcelResourceDto>();
                
                foreach (ExcelResourceDto dto in objectsList)
                {
                    Console.WriteLine(dto.Title);
                }
            }
        }
    }
}
