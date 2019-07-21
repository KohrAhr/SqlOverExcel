using System;
using System.Data;
using ExcelWorkbookSplitter.Functions;
using ExcelWorkbookSplitter.Core;
using System.Threading.Tasks;
using System.Threading;

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
                using (ExcelCore excelIn = new ExcelCore(appData.inFile, appData.acever))
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
                            CancellationTokenSource tokenSource = new CancellationTokenSource();

                            try
                            {
                                Console.WriteLine("\nSQL Query execution running at: {0}", DateTime.Now.ToLongTimeString());

                                // Start progress...
                                CancellationToken ct = tokenSource.Token;

                                const int CONST_DELAY = 500;
                                string[] values = { "|", "/", "-", "\\" };

                                Task progress = Task.Run(() =>
                                {
                                    while (true)
                                    {
                                        //                                        coreFunctions.ShowProgressInConsole();

                                        foreach (string s in values)
                                        {
                                            Console.Write("\r{0}", s);
                                            Thread.Sleep(CONST_DELAY);

                                            if (ct.IsCancellationRequested)
                                            {
                                                Console.Write("\r ");
                                                ct.ThrowIfCancellationRequested();
                                            }
                                        }
                                    }
                                }, tokenSource.Token);

                                // Start query
                                excelIn.RunSql(appData.query, ref queryResult);

                                Console.WriteLine("\nSQL Query execution completed at: {0}", DateTime.Now.ToLongTimeString());
                         
                                if (appData.resultToFile)
                                {
                                    coreFunctions.SaveResultToExcelFile(appData.outFile, queryResult);
                                }
                                else
                                {
                                    coreFunctions.DisplayResult(queryResult);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Cancel progress
                                coreFunctions.StopProgress(tokenSource);

                                Console.WriteLine(
                                    "\nThe an error has been occured during executing the SQL query: \nSQL Query: \"{0}\"\nFile: \"{1}\"\nError message: {2}",
                                    appData.query,
                                    appData.inFile,
                                    ex.Message
                                );
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("\nThe an error has been occured during accessing Excel file \"{0}\"", appData.inFile);
                    }
                }
            }

            coreFunctions.ByeMessage();
        }
    }
}
