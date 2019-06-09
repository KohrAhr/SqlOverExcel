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
            ExcelFile excelFile = new ExcelFunctions().OpenFile(@"E:\Temp\GitHubUnsorted\Wpf.ProductsFromExcelToXml\!Data\VAIRUMTIRDZNIECIBA_matraci_TEEN_2018.xlsx");

            List<string> worksheets = new ExcelFunctions().GetWorksheets(excelFile);

            ExcelObject.Worksheet worksheet = new ExcelFunctions().GetWorksheet(excelFile, "KIDS");

            //            using ()
            {
                //                ExcelObject.Worksheet worksheet = excelFile.sheet.Worksheets["KIDS"];

                foreach (String name in worksheets)
                {
                    Console.WriteLine(name);
                }

                //                Console.WriteLine("Total lines: {0}", x.ToString());
            }

            Console.ReadKey();
        }
    }
}
