using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace SalesAnalysis.TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start");

            FileInfo fileInfo = new FileInfo(@"C:\Pawlak\monika.xlsx");

            using (ExcelPackage pck = new ExcelPackage(fileInfo))
            {
                var ws = pck.Workbook.Worksheets.Add("Content");
                ws.View.ShowGridLines = false;


            }

                Console.WriteLine("Stop");
            Console.ReadLine();
        }
    }
}
