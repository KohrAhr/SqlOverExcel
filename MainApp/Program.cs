using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelWorkbookSplitter.Functions;
using ExcelObject = Microsoft.Office.Interop.Excel;
using Lib.Strings;
using Lib.System;

namespace ExcelWorkbookSplitter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Command line usage sample:
            // APP.EXE in="EXCEL FILE NAME" [out="EXCEL FILE NAME"] [query="SQL Query"]

            // Parse command line params


            //


            string inFile = @"C:\Temp\Excel\test.xlsx";
            string outFile = "";
//            outFile = @"C:\Temp\Excel\NewExcelFile.xlsx";
            string query = "";
//            query = "select count(field1) as e1 from [data$]";
            query = "select count(field1) as e1 from [data$]";

            bool info = String.IsNullOrEmpty(query);
            bool resultToFile = !String.IsNullOrEmpty(outFile);

            // Open file and ...
            using (ExcelCore excelIn = new ExcelCore(inFile))
            {
                if (excelIn.IsInitialized())
                {
                    if (info)
                    {
                        new CoreFunctions().DisplayWorksheetInfo(excelIn);
                    }
                    else
                    {
                        DataTable queryResult = new DataTable();
                        if (excelIn.RunSql(query, ref queryResult))
                        {
                            if (resultToFile)
                            {
                                // Option 1

                                // Save result to new file
                                using (ExcelCore excelOut = new ExcelCore())
                                {
                                    excelOut.NewFile(outFile);
                                    if (excelOut.IsInitialized())
                                    {
                                        if (excelOut.NewSheet("RESULT", WorksheerOrder.woFirst))
                                        {
                                            // Delete default worksheet
                                            excelOut.DeleteSheet("Sheet1");

                                            excelOut.PopulateData("RESULT", queryResult);

                                            excelOut.SaveFile();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                // Option 2

                                new CoreFunctions().DisplayResult(queryResult);
                            }
                        }
                        else
                        {
                            Console.WriteLine("The an error has been occured during executing the sql query");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("The an error has been occured during accessing Excel file");
                }

                Console.ReadKey();

                return;
            }
        }
    }
}
