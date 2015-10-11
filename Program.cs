using AQ1Report.BC;
using System;

namespace AQ1Report
{
    class Program
    {
        static void Main(string[] args)
        {
            String inputFileName, outputFileName;

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("PROGRAM INFORMATION:");
            Console.ResetColor();
            Console.WriteLine("This program takes 1 excel file (called AQ1 Report) as input, extracts information from it and generates a new excel file based on the extracted information.");
            Console.WriteLine("INPUT: AQ1 Report (xlsx file) containing multiple worksheets.");
            Console.WriteLine("OUTPUT: Excel (xlsx) file with a single worksheet containing all the information extracted from the input file.");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("INPUT PARAMETERS:");
            Console.ResetColor(); 
            Console.WriteLine("Input Parameter 1: Please provide complete path to input excel file (AQ1 Report):");
            inputFileName = Console.ReadLine();
            Console.WriteLine("");
            Console.WriteLine("Input Parameter 2: Please provide complete path to output excel file:");
            outputFileName = Console.ReadLine();
            Console.WriteLine("");
            Console.WriteLine("");

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("PROCESSING:");
            Console.ResetColor();
            clsAQ1ReportBC.ProcessAQ1Excel(inputFileName, outputFileName);
        }
    }
}
