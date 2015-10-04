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
            String fileNameTest = @"C:\wip\Didi\Test.xlsx";
            String fileNameAQ1 = @"C:\wip\Didi\AQ 1 317 JUNE 15-ZamZar.xlsx";
            String outputFileName = @"C:\wip\Didi\AQ1-Consolidated.xlsx";

            // Comment one of the following lines to test the method separately.
            //Console.WriteLine("ReadExcelFileDOM....");
            //clsOpenXmlBC.ReadExcelFileDOM(fileName);    // DOM
            //Console.WriteLine("");
            //Console.WriteLine("");
            //Console.WriteLine("ReadExcelFileSAX....");
            //clsOpenXmlBC.ReadExcelFileSAX(fileNameTest);    // SAX

            //Console.WriteLine("");
            //Console.WriteLine("");
            //Console.WriteLine("ReadExcelFileStackTrace....");
            //clsOpenXmlBC.ReadExcelFileStackTrace(fileNameTest);

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("ProcessAQ1Excel....");
            clsAQ1ReportBC.ProcessAQ1Excel(fileNameAQ1, outputFileName);

            //Console.WriteLine("");
            //Console.WriteLine("");
            //Console.WriteLine("ExcelBasics....");
            //clsOpenXmlBC.ExcelBasics(fileNameAQ1);
        }
    }
}
