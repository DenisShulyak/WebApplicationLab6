using System;
using System.Collections.Generic;
using System.Text;

namespace WebApplicationLab6.Objects
{
    public class Report
    {
       public List<ReportRow> Rows { get; set; } = new List<ReportRow>();

       public DateTime Date = DateTime.Now;
    }
}
