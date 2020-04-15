using System;
using System.Collections.Generic;
using System.Text;

namespace ReadFromExcel.Models
{
        public class All
        {
            public int Id { get; set; }

            public DateTime Date { get; set; }
            public int Cases { get; set; }
            public int Deaths { get; set; }
            public int Recovered { get; set; }
        }
}
