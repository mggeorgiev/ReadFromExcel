using System;
using System.Collections.Generic;
using ReadFromExcel.Infrastructure;
using ReadFromExcel.Models;

namespace ReadFromExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            StepOne();
            StepTwo();
        }


        private static void StepOne()
        {
            Console.WriteLine("Step One: load data");

            IEnumerable<All> data = Data();
            Excel.SaveToExcel(data);
        }
        private static void StepTwo()
        {
            Console.WriteLine("Step two: read data");
            Excel.ReadFromExcelFile();
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
