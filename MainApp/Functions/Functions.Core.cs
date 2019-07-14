using ExcelWorkbookSplitter.Core;
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
            Console.WriteLine("======================================");
            Console.WriteLine("SqlOverExcel. Version 0.1 from 11/Jul/2019");
            Console.WriteLine("======================================");
            Console.WriteLine("Run SQL query over Excel file\n");
            Console.WriteLine("Usage: SqlOverQuery.exe -in=\"EXCEL FILE NAME\" [-out=\"EXCEL FILE NAME\"] -query=\"SQL Query\"]");
            Console.WriteLine("\nOptions:");
            Console.WriteLine("\t-in        \tSource Excel file");
            Console.WriteLine("\t-out       \tOutput file. If not provided, than result of query execution will be displayed in console");
            Console.WriteLine("\t-query     \tSQL Query to run. SQL query compactible with MS Access");
            Console.WriteLine("\nSQL Query samples:");
            Console.WriteLine("\tselect count(field1) as e1 from [Worksheet1]");
            Console.WriteLine("\tselect count(field1) as e1, max(field2) as e2, min(field3) as e3 from [Worksheet1]");
            Console.WriteLine("\nAllowed file types:");
            Console.WriteLine("\tXLSX -- Excel Workbook (..Office 2016)");
            Console.WriteLine("\tXLSM -- Excel Macro-Enabled Workbook");
            Console.WriteLine("\tXLSB -- Excel Binary Workbook");
            Console.WriteLine("\tXLS  -- Excel 97-2003 Workbook");
        }

        public void ByeMessage()
        {
            Console.WriteLine("Press any key for conitnue...");
            Console.ReadKey();
            return;
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

        public void SaveResultToExcelFile(string toFile, DataTable data)
        {
            // Save result to new file
            using (ExcelCore excelOut = new ExcelCore())
            {
                excelOut.NewFile(toFile);
                if (excelOut.IsInitialized())
                {
                    if (excelOut.NewSheet("RESULT", WorksheerOrder.woFirst))
                    {
                        // Delete default worksheet
                        excelOut.DeleteSheet("Sheet1");

                        excelOut.PopulateData("RESULT", data);

                        excelOut.SaveFile();
                    }
                }
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

        public AppData ParseCommandLineParams(string[] arguments)
        {
            AppData result = new AppData();

            foreach(string s in arguments)
            {
                // Get key and value
                string[] input = s.Split(new Char[] { '=' });

                // Is key=value pair?
                if (input.Length != 2)
                {
                    continue;
                }

                // Key in uppercase
                string key = input[0].ToLower();

                if (key == "-in")
                {
                    result.inFile = input[1];
                }
                else
                if (key == "-out")
                {
                    result.outFile = input[1];
                }
                else
                if (key == "-query")
                {
                    result.query = input[1];
                }
            }
            
            return result;
        }

        /// <summary>
        ///     Simple validation of command line params
        /// </summary>
        /// <param name="appData"></param>
        /// <returns></returns>
        public AppData ValidateCommandLineParams(AppData appData)
        {
            AppData result = appData;

            if (!String.IsNullOrEmpty(result.outFile) && String.IsNullOrEmpty(result.query))
            {
                Console.WriteLine("You didn't provide SQL query to run. Output fill will be not created");
                result.outFile = "";
            }

            return result;
        }
    }
}
