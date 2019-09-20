using System;
using System.Data;
using SqlOverExcel.Functions;
using SqlOverExcel.Core;
using System.Threading.Tasks;
using System.Threading;
using Lib.Excel;

namespace SqlOverExcel
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
                try
                {
                    using (ExcelCore excelIn = new ExcelCore(appData.inFile, appData.acever, appData.useHdr == "Y" ? true : false))
                    {
                        if (appData.showInfo)
                        {
                            coreFunctions.DisplayWorksheetInfo(excelIn);
                        }
                        else
                        {
                            DataTable queryResult = new DataTable();
                            CancellationTokenSource tokenSource = new CancellationTokenSource();

                            try
                            {
                                Console.WriteLine("\nSQL Query execution running at: {0}", DateTime.Now.ToString());

                                // Start progress...
                                coreFunctions.ShowProgressInConsole(tokenSource.Token);

                                // Start query
                                excelIn.RunSql(appData.query, ref queryResult);

                                // Cancel progress
                                coreFunctions.StopProgress(tokenSource);
                                tokenSource = null;

                                Console.WriteLine("\nSQL Query execution completed at: {0}", DateTime.Now.ToString());

                                if (appData.resultToFile)
                                {
                                    excelIn.SaveResultToExcelFile(appData.outFile, queryResult);
                                }
                                else
                                {
                                    coreFunctions.DisplayResult(queryResult);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Cancel progress
                                if (tokenSource != null)
                                {
                                    coreFunctions.StopProgress(tokenSource);
                                }

                                Console.WriteLine(
                                    "\nThe an error has been occured during executing the SQL query: \nSQL Query: \"{0}\"\nFile: \"{1}\"\nError message: {2}",
                                    appData.query,
                                    appData.inFile,
                                    ex.Message
                                );
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("\nThe an error has been occured during accessing Excel file \"{0}\"\nError message: {1}", appData.inFile, ex.Message);
                }
            }

            coreFunctions.ByeMessage();
        }
    }
}
