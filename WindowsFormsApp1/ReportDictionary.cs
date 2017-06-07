using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class ReportDictionary : Dictionary<String, String>
    {
        public Report BuildReport()
        {
            Report report = new Report();
            String field;
            this.TryGetValue("Hello", out field);
            report.ReportName = field;
            return report;
        }
    }
}
