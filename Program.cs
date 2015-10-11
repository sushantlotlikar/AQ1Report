using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AQ1Report.BC;

namespace AQ1Report
{
    class Program
    {
        static void Main(string[] args)
        {
            String fileNameAQ1 = @"C:\wip\Didi\AQ 1 317 JUNE 15-ZamZar.xlsx";
            String outputFileName = @"C:\wip\Didi\AQ1-Consolidated.xlsx";

            Console.WriteLine("ProcessAQ1Excel....");
            clsAQ1ReportBC.ProcessAQ1Excel(fileNameAQ1, outputFileName);
        }
    }
}
