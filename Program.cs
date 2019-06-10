using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelWorkbookSplitter.Functions;
using ExcelObject = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookSplitter
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelCore excelCore = new ExcelCore(@"C:\Temp\Test.xlsx"))
            {
                List<string> worksheets = excelCore.GetWorksheets();

//                foreach (String name in excelCore.excelFile.Worksheets)
//               {
 //                  Console.WriteLine(name);
    //            }
            }

            //ExcelObject.Worksheet worksheet = new ExcelFunctions().GetWorksheet(excelFile, 1);

            //            using ()
            //{
            //    //                ExcelObject.Worksheet worksheet = excelFile.sheet.Worksheets["KIDS"];

            //    foreach (String name in worksheets)
            //    {
            //        Console.WriteLine(name);
            //    }

            //    //                Console.WriteLine("Total lines: {0}", x.ToString());
            //}

            Console.WriteLine("Done!");
            Console.ReadKey();
        }
    }
}
