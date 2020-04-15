using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ReadFromExcel.Models;

namespace ReadFromExcel.Infrastructure
{
    class csv
    {
        public static void SaveToCSV(IEnumerable<All> all)
        {

            var builder = new StringBuilder();
            builder.AppendLine("Id,Date,Cases,Recovered,Deaths");

            foreach (var item in all)
            {
                builder.AppendLine($"{item.Id},{item.Date},{item.Cases},{item.Recovered},{item.Deaths}");
            }

            //File.AppendAllTextAsync(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "all" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
        }

    }
}
