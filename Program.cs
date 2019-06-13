using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
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
            using (ExcelCore excelCore = new ExcelCore(@"C:\Temp\Excel\test.xlsx"))
            {
                if (excelCore.IsInitialized())
                {
                    // Display common information about Excel worksheets

                    List<string> worksheets = excelCore.GetWorksheets();

                    Console.WriteLine("List of available worksheets in file \"{0}\":", excelCore.FileName);
                    foreach (String name in worksheets)
                    {
                        Console.WriteLine("\t\"{0}\"", name);

                        ExcelObject.Worksheet worksheet = excelCore.GetWorksheet(name);

                        Console.WriteLine("\tLast row with data: {0}; Last column with data: {1}\n", 
                            excelCore.GetCountOfRows(worksheet), 
                            excelCore.GetCountOfCols(worksheet)
                        );
                    }

                    // Get worksheet

                    DataTable dataTables = new DataTable();
                    DataTable dataTable = new DataTable();

                    if (excelCore.GetTables(ref dataTables))
                    {
                        Console.WriteLine("Count of available tables (worksheet): {0}", dataTables.Rows.Count);

                        // ... from first worksheet (index = 0)
                        // name == actual name + $
                        if (excelCore.GetListData("dAtA$", ref dataTable))
                        {
                            // 
                            DataTable testData = new DataTable();
                            if (excelCore.RunSql("select count(Field1) as E1 from [dAtA$]", ref testData))
                            {


                            }
                            else
                            {
                                Console.WriteLine("Error occured during executing SQL query");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error occured during obtaining data from Table");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Error occured during obtaining list of tables");
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
