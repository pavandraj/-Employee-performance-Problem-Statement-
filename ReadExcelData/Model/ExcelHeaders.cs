using System;
using System.Collections.Generic;
using System.Text;

namespace ReadExcelData.Model
{
    public class ExcelHeaders
    {
        public DateTime Date { get; set; }
        public string ProjectName { get; set; }
        public Double Hours { get; set; }
        public string Owner { get; set; }
        public string Team { get; set; }
        public string BillingStatus { get; set; }
    }
}
