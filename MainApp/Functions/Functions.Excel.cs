using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using ExcelObject = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookSplitter.Functions
{
    /// <summary>
    ///     Main class
    /// </summary>
    public class ExcelCore : ExcelFile, IDisposable
    {
        /// <summary>
        ///     Default constructor
        /// </summary>
        public ExcelCore()
        {
        }

        /// <summary>
        ///     Constructor with file name
        /// </summary>
        /// <param name="FileName">
        ///     Excel file name
        /// </param>
        public ExcelCore(String FileName) : this()
        {
            OpenFile(FileName);
        }

        /// <summary>
        ///     Destructor
        /// </summary>
        public void Dispose()
        {
            CloseFile();
        }

        /// <summary>
        ///     Create instance on Excel Application.
        ///     Instance Not visible and Does Not display alerts
        /// </summary>
        /// <returns>
        ///     Excel instance
        /// </returns>
        private ExcelObject.Application CreateExcelInstance()
        {
            ExcelObject.Application excelInstance = new ExcelObject.Application();
            excelInstance.Visible = false;
            excelInstance.DisplayAlerts = false;

            return excelInstance;
        }

        /// <summary>
        ///     Open existing Excel file
        /// </summary>
        /// <param name="file">
        ///     Excel file name and optional path
        /// </param>
        public string OpenFile(string file)
        {
            string result = "";
            FileName = file;

            try
            {
                ExcelApp = CreateExcelInstance();
                Books = ExcelApp.Workbooks;
                Sheet = Books.Open(file);

                if (Sheet == null)
                {
                    CloseFile();
                }
            }
#pragma warning disable 168
            catch (Exception ex)
#pragma warning restore 168
            {
                result = ex.Message.ToString();
                CloseFile();
            }

            return result;
        }

        /// <summary>
        ///     Create new Excel file
        /// </summary>
        /// <param name="file">
        ///     Excel file name and optional path
        /// </param>
        public void NewFile(string file)
        {
            FileName = file;

            try
            {
                ExcelApp = CreateExcelInstance();
                Books = ExcelApp.Workbooks;
                Sheet = Books.Add();
            }
            catch
            {
                CloseFile();
            }
        }

        /// <summary>
        ///     Create new worksheet 
        /// </summary>
        /// <param name="sheetName">
        ///     Worksheet name (optional)
        /// </param>
        /// <returns>
        ///     True if operation completed successfully
        /// </returns>
        public bool NewSheet(string sheetName = "", WorksheerOrder workSheetOrder = WorksheerOrder.woFirst)
        {
            ExcelObject.Worksheet worksheet = null;
            try
            {
                ExcelObject.Sheets xlSheets = ExcelApp.Worksheets;

                int position = 1;

                if (workSheetOrder == WorksheerOrder.woLast)
                {
                    position = xlSheets.Count;
                }

                worksheet = xlSheets.Add(xlSheets[position]);
                worksheet.Name = sheetName;
            }
#pragma warning disable 168
            catch (Exception ex)
#pragma warning restore 168
            {
                worksheet = null;
            }

            return worksheet != null;
        }

        /// <summary>
        ///     Find Worksheet Id by Worksheet Name
        ///     <para>Worksheet Id starting from 1</para>
        /// </summary>
        /// <param name="worksheetName">
        ///     Worksheet name
        /// </param>
        /// <returns>
        ///     0 -- Worksheet not found
        /// </returns>
        public int FindWorksheet(string worksheetName)
        {
            int result = 0;

            ExcelObject.Worksheet worksheet = Sheet.Worksheets[worksheetName];

            if (worksheet != null)
            {
                result = worksheet.Index;
            }

            return result;
        }

        /// <summary>
        ///     Delete worksheet by Worksheet Id
        ///     <para>You cannot delete the last one worksheet</para>
        /// </summary>
        /// <param name="sheetIndex">
        ///     Worksheet Id. Index starting from 1
        /// </param>
        /// <returns>
        ///     True if operation completed successfully
        /// </returns>
        public bool DeleteSheet(int sheetIndex)
        {
            bool result = true;

            try
            {
                ExcelApp.Sheets[sheetIndex].Delete();
            }
            catch
            {
                result = false;
            }
            
            return result;
        }

        /// <summary>
        ///     Delete worksheet by Worksheet name
        ///     <para>You cannot delete last worksheet</para>
        /// </summary>
        /// <param name="sheetIndex">
        ///     Worksheet Id. Index starting from 1
        /// </param>
        /// <returns>
        ///     True if operation completed successfully
        /// </returns>
        public bool DeleteSheet(string sheetName)
        {
            return DeleteSheet(FindWorksheet(sheetName));
        }

        public bool PopulateData(string sheetName, DataTable data)
        {
            bool result = true;

            int sheetIndex = FindWorksheet(sheetName);

            if (sheetIndex > 0)
            {
                try
                {
                    int y = 1;
                    foreach (DataRow dataRow in data.Rows)
                    {
                        int x = 1;
                        foreach (object item in dataRow.ItemArray)
                        {
                            ExcelApp.Sheets[sheetIndex].Cells[y, x++] = item.ToString();
                        }
                        y++;
                    }
                }
                catch (Exception ex)
                {
                    result = false;

                    throw new Exception(ex.Message);
                }
            }
            else
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        ///     Save Excel file As
        /// </summary>
        /// <param name="file">
        ///     Excel file name and optional path
        /// </param>
        /// <returts>
        ///     True if operation completed successfully
        /// </returts>
        public bool SaveFile(string file)
        {
            bool result = true;

            try
            {
                Sheet?.SaveAs(file,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
                    ExcelObject.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                );
            }
            catch
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        ///     Save Excel file    
        /// </summary>
        public void SaveFile()
        {
            SaveFile(FileName);
        }

        /// <summary>
        ///     Close Excel file and release all Excel objects
        /// </summary>
        public void CloseFile()
        {
            if (Worksheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Worksheet);
            }
            Sheet?.Close(false);
            Books?.Close();
            if (Sheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Sheet);
            }
            if (Books != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Books);
                Books = null;
            }
            ExcelApp?.Quit();
            if (ExcelApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                ExcelApp = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        ///     Get worksheet by Worksheet name
        /// </summary>
        /// <param name="worksheetName">
        ///     Worksheet name
        /// </param>
        /// <returns>
        ///     Worksheet
        /// </returns>
        public ExcelObject.Worksheet GetWorksheet(String worksheetName)
        {
            return Sheet.Worksheets[worksheetName];
        }

        /// <summary>
        ///     Get worksheet by Worksheet Id
        /// </summary>
        /// <param name="worksheetName">
        ///     Worksheet Id
        /// </param>
        /// <returns>
        ///     Worksheet
        /// </returns>
        public ExcelObject.Worksheet GetWorksheet(int worksheetId)
        {
            return Sheet.Worksheets[worksheetId];
        }

        /// <summary>
        ///     Get list of all Worksheets
        /// </summary>
        /// <returns>
        ///     List of Worksheets
        /// </returns>
        public List<string> GetWorksheets()
        {
            List<string> result = new List<string>();
            foreach (ExcelObject.Worksheet ws in Sheet.Worksheets)
            {
                result.Add(ws.Name?.ToString());
            }
            return result;
        }

        /// <summary>
        ///     Get cell value from specific Worksheet
        /// </summary>
        /// <param name="excelWorksheet">
        ///     Worksheet
        /// </param>
        /// <param name="row">
        ///     Row No
        /// </param>
        /// <param name="col">
        ///     Col No
        /// </param>
        /// <returns>
        ///     Value. If value is Null will be returned Empty string
        /// </returns>
        public String GetCellValue(ExcelObject.Worksheet excelWorksheet, int row, int col)
        {
            return excelWorksheet.Cells[row, col].Value2?.ToString();
        }

        /// <summary>
        ///     Get maximum allowed count of rows in Worksheet
        /// </summary>
        /// <param name="excelWorksheet">
        ///     Worksheet to check
        /// </param>
        /// <returns>
        ///     Numeric value -- maximum allowed count of rows in specific Worksheet
        /// </returns>
        public int GetMaxCountOfRows(ExcelObject.Worksheet excelWorksheet)
        {
            return excelWorksheet.Rows.Count;
        }

        /// <summary>
        ///     Get maximum allowed count of columns in Worksheet
        /// </summary>
        /// <param name="excelWorksheet">
        ///     Worksheet to check
        /// </param>
        /// <returns>
        ///     Numeric value -- maximum allowed count of columns in specific Worksheet
        /// </returns>
        public int GetMaxCountOfCols(ExcelObject.Worksheet excelWorksheet)
        {
            return excelWorksheet.Columns.Count;
        }

        /// <summary>
        ///     Find last row or column in worksheet
        /// </summary>
        /// <param name="excelWorksheet">
        ///     Worksheet for analuze
        /// </param>
        /// <param name="searchOrder">
        ///     By Row or by Column
        /// </param>
        /// <returns>
        ///     Cell
        /// </returns>
        private dynamic GetCount(ExcelObject.Worksheet excelWorksheet, ExcelObject.XlSearchOrder searchOrder)
        {
            return excelWorksheet.Cells.Find(
                "*",
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                searchOrder,
                ExcelObject.XlSearchDirection.xlPrevious,
                false,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value
            );
        }

        /// <summary>
        ///     Get count of actually filled rows in Worksheet
        /// </summary>
        /// <param name="excelWorksheet">
        ///     Worksheet to check
        /// </param>
        /// <returns></returns>
        public int GetCountOfRows(ExcelObject.Worksheet excelWorksheet)
        {
            return GetCount(excelWorksheet, ExcelObject.XlSearchOrder.xlByRows).Row;
        }

        /// <summary>
        ///     Get count of maximum actually filled columns in Worksheet
        /// </summary>
        /// <param name="excelWorksheet"></param>
        /// <returns></returns>
        public int GetCountOfCols(ExcelObject.Worksheet excelWorksheet)
        {
            return GetCount(excelWorksheet, ExcelObject.XlSearchOrder.xlByColumns).Column;
        }

        /// <summary>
        ///     Get worksheet as Table
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public bool GetTables(ref DataTable dt)
        {
            OleDbConnection oConn = null;
            try
            {
                String sConnString = BuildConnectionString();

                oConn = new OleDbConnection(sConnString);
                oConn.Open();

                dt = oConn.GetSchema("Tables");

                return dt.Rows.Count > 0;
            }
#pragma warning disable 168
            catch (Exception ex)
#pragma warning restore 168
            {
                return false;
            }
            finally
            {
                oConn.Close();
                oConn.Dispose();
            }
        }

        /// <summary>
        ///     Run SQL Query over Excel file
        /// </summary>
        /// <param name="sql">
        ///     SQL Query to run
        /// </param>
        /// <param name="dataTable">
        ///     DataSet with result
        /// </param>
        /// <returns>
        ///     True if operation completed successfully
        /// </returns>
        public void RunSql(string sql, ref DataTable dataTable)
        {
            OleDbConnection oConn = null;
            OleDbCommand oComm = null;
            OleDbDataReader oRdr = null;
            try
            {
                try
                {
                    String sConnString = BuildConnectionString();

                    oConn = new OleDbConnection(sConnString);
                    oConn.Open();

                    String sCommand = sql;
                    oComm = new OleDbCommand(sCommand, oConn);
                    oRdr = oComm.ExecuteReader();

                    dataTable.Load(oRdr);
                }
                finally
                {
                    oRdr?.Close();
                    oRdr = null;
                    oComm?.Dispose();
                    oConn.Close();
                    oConn.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
                //                Console.WriteLine(ex.Message);
            }
        }

        public Action<string, int, int> TIterateAction;
        public class IterateAction
        {
            public Action<string, int, int> action;
        }

        public void IterateOverData(DataTable dataTable, IteratorEvent action)
        {
            int y = 1;
            foreach (DataRow dataRow in dataTable.Rows)
            {
                int x = 1;
                foreach (object item in dataRow.ItemArray)
                {
                    action?.Invoke(item.ToString(), x++, y++);
                }
            }
        }

        /// <summary>
        ///     Build connection string for current Excel file
        /// </summary>
        /// <returns></returns>
        public string BuildConnectionString()
        {
            const string connectionStringTemplate = @"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=Yes';";

            return String.Format(connectionStringTemplate, FileName);
        }
    }
}
