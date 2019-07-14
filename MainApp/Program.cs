using System;
using System.Data;
using ExcelWorkbookSplitter.Functions;
using ExcelWorkbookSplitter.Core;

namespace ExcelWorkbookSplitter
{
    class Program
    {
        static void Main(string[] args)
        {
            CoreFunctions coreFunctions = new CoreFunctions();

            // Parse command line params
            AppData appData = coreFunctions.ParseCommandLineParams(args);
            appData = coreFunctions.ValidateCommandLineParams(appData);

            if (appData.showHelp)
            {
                coreFunctions.ShowHelp();
            }
            else
            {
                // Open file and ...
                using (ExcelCore excelIn = new ExcelCore(appData.inFile))
                {
                    if (excelIn.IsInitialized())
                    {
                        if (appData.showInfo)
                        {
                            coreFunctions.DisplayWorksheetInfo(excelIn);
                        }
                        else
                        {
                            DataTable queryResult = new DataTable();
                            if (excelIn.RunSql(appData.query, ref queryResult))
                            {
                                if (appData.resultToFile)
                                {
                                    coreFunctions.SaveResultToExcelFile(appData.outFile, queryResult);
                                }
                                else
                                {
                                    coreFunctions.DisplayResult(queryResult);
                                }
                            }
                            else
                            {
                                Console.WriteLine("The an error has been occured during executing the SQL query");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("The an error has been occured during accessing Excel file \"{0}\"", appData.inFile);
                    }
                }
            }

            Console.ReadKey();
            return;
        }
    }
}
