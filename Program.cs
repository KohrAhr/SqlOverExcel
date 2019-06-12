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
            using (ExcelCore excelCore = new ExcelCore(@"C:\Temp\Excel\Book1.xlsx"))
            {
                if (excelCore.IsInitialized())
                {
                    List<string> worksheets = excelCore.GetWorksheets();

                    Console.WriteLine("List of available worksheets:");
                    foreach (String name in worksheets)
                    {
                        Console.WriteLine("\t\"{0}\"", name);

                        ExcelObject.Worksheet worksheet = excelCore.GetWorksheet(name);

                        Console.WriteLine("\tMax rows: {0}; Max cols: {1}", excelCore.GetMaxCountOfRows(worksheet), excelCore.GetMaxCountOfCols(worksheet));
                        Console.WriteLine("\tActual rows: {0}; Actual cols: {1}", excelCore.GetCountOfRows(worksheet), excelCore.GetCountOfCols(worksheet));
                    }
                }
                else
                {
                    Console.WriteLine("Requested file cannot be opened");
                }
            }

            Console.WriteLine("Done!");
            Console.ReadKey();
        }
    }
}
