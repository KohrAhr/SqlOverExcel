using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelObject = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookSplitter.Functions
{
    public class CoreFunctions
    {
        public void ShowHelp()
        {
            Console.WriteLine("\n\n======================================");
            Console.WriteLine("SqlOverQuery. Version 0.1 from 11/Jul/2019");
            Console.WriteLine("======================================\n");
            Console.WriteLine("Usage: SqlOverQuery.exe -in=\"EXCEL FILE NAME\" [-out=\"EXCEL FILE NAME\"] -query=\"SQL Query\"]");
            Console.WriteLine("\nOptions:");
            Console.WriteLine("\t-in        \tSource Excel file");
            Console.WriteLine("\t-out       \tOutput file. If not provided, than result of query execution will be displayed in console");
            Console.WriteLine("\t-query     \tSQL Query to run. SQL query compactible with MS Access");
            Console.WriteLine("\nAllowed file types:");
            Console.WriteLine("\n\tXLSX -- Excel Workbook (..Office 2016)");
            Console.WriteLine("\n\tXLSM -- Excel Macro-Enabled Workbook");
            Console.WriteLine("\n\tXLSB -- Excel Binary Workbook");
            Console.WriteLine("\n\tXLS  -- Excel 97-2003 Workbook");
        }

        /// <summary>
        ///     Verbose output -- show result from datatable
        /// </summary>
        /// <param name="data"></param>
        public void DisplayResult(DataTable data)
        {
            using (ExcelCore excel = new ExcelCore())
            {
                excel.IterateOverData(data,
                    delegate (string value, int x, int y)
                    {
                        if (x == 1)
                        {
                            Console.WriteLine();
                        }
                        Console.Write(value + "\t");
                    }
                );
            }
        }

        /// <summary>
        ///     Display common information about Excel worksheets
        /// </summary>
        /// <param name="excel"></param>
        public void DisplayWorksheetInfo(ExcelCore excel)
        {
            List<string> worksheets = excel.GetWorksheets();

            Console.WriteLine("List of available worksheets in file \"{0}\":", excel.FileName);
            foreach (String name in worksheets)
            {
                Console.WriteLine("\t\"{0}\"", name);

                ExcelObject.Worksheet worksheet = excel.GetWorksheet(name);

                Console.WriteLine("\tLast row with data: {0}; Last column with data: {1}\n",
                    excel.GetCountOfRows(worksheet),
                    excel.GetCountOfCols(worksheet)
                );
            }
        }
    }
}
