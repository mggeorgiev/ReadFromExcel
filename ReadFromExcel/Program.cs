using System;
using System.Collections.Generic;
using ReadFromExcel.Infrastructure;
using ReadFromExcel.Models;

namespace ReadFromExcel
{
    class Program
    {
        Console.OutputEncoding = Encoding.UTF8;
        static void Main(string[] args)
        {
            IEnumerable<All> data = Data();
            StepOne(data);
            PrintTestData(data);
            StepTwo();
        }


        private static void StepOne(IEnumerable<All> data)
        {
            Console.WriteLine("Step One: load data");

            Excel.SaveToExcel(data);
        }
        private static void StepTwo()
        {
            Console.WriteLine("Step two: read data");
            Excel.ReadFromExcelFile();
        }

        private static void PrintTestData(IEnumerable<All> data)
        {
            foreach(var item in data)
            {

                Console.WriteLine("{0},{1},{2},{3},{4}", item.Id, item.Date, item.Cases, item.Recovered, item.Deaths);
            }
        }

        private static IEnumerable<All> Data()
        {
            return new List<All>
            {
                new All{Id = 1, Cases= 1, Date = DateTime.Now, Deaths= 0, Recovered=0},
                new All{Id = 2, Cases= 2, Date = DateTime.Now, Deaths= 1, Recovered=1},
                new All{Id = 3, Cases= 3, Date = DateTime.Now, Deaths= 1, Recovered=2}
            };
        }
    }
}
