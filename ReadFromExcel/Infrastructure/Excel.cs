using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using ReadFromExcel.Models;

namespace ReadFromExcel.Infrastructure
{
    class Excel
    {
        public static void SaveToExcel(IEnumerable<All> all)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("All");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Date";
                worksheet.Cell(currentRow, 3).Value = "Cases";
                worksheet.Cell(currentRow, 4).Value = "Recovered";
                worksheet.Cell(currentRow, 5).Value = "Deaths";

                foreach (var item in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = item.Id;
                    worksheet.Cell(currentRow, 2).Value = item.Date;
                    worksheet.Cell(currentRow, 3).Value = item.Cases;
                    worksheet.Cell(currentRow, 4).Value = item.Recovered;
                    worksheet.Cell(currentRow, 5).Value = item.Deaths;
                }

                var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\TestData.xlsx");
                workbook.SaveAs(fileName);

                //using (var stream = new MemoryStream())
                //{
                    //workbook.SaveAs( (stream);
                    //var content = stream.ToArray();
                    //stream.Write

                    //return File(
                    //    content,
                    //    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    //    "all-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx");
                //}
            }
        }

        public static void ReadFromExcelFile()
        {
            var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\TestData.xlsx");
            var workbook = new XLWorkbook(fileName);
            var ws1 = workbook.Worksheet(1);
            int iRow = 1;
            while (!ws1.Cell(iRow,1).IsEmpty())
            {
                var row = "";
                int iColumn = 1;
                while(!ws1.Cell(iRow,iColumn).IsEmpty())
                {
                    row = row + ws1.Cell(iRow, iColumn).Value.ToString() + ",";
                    iColumn++;
                }
                Console.WriteLine(row);
                iRow++;
            }
        }
    }   
}
