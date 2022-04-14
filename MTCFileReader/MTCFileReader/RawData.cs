using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MTCFileReader
{
    public class RawData
    {
        public string store_code { get; set; }
        public string assing_code { get; set; }
        public DateTime process_date { get; set; }
        public DateTime register_date { get; set; }
        public DateTime paid_date { get; set; }
        public string concept { get; set; }
        public decimal total { get; set; }
    }
}
