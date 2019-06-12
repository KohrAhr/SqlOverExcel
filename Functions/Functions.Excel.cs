using System;
using System.Collections.Generic;
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
        ///     Open existing Excel file
        /// </summary>
        /// <param name="file">
        ///     Excel file name and optional path
        /// </param>
        public void OpenFile(string file)
        {
            FileName = file;

            try
            {
                ExcelApp = new ExcelObject.Application();
                ExcelApp.Visible = false;
                Books = ExcelApp.Workbooks;
                Sheet = Books.Open(file);

                if (Sheet == null)
                {
                    CloseFile();
                }
            }
            catch
            {
                CloseFile();
            }
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

        public static string GetString(ExcelObject.Worksheet worksheet, int row, int col)
        {
            string result = "";

            ExcelObject.Range range = worksheet.Cells[row, col];
            try
            {
                result = range.Value2?.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            }

            return result;
        }

        //public IEnumerable<string> GetHeaderFromWorksheet(ExcelObject.Worksheet excelWorksheet, int headerEnd)
        //{
        //    for (int intPosition = 0; intPosition < headerEnd; intPosition++)
        //    {
        //        yield return GetLineFromWorksheet(excelWorksheet, intPosition);
        //    }
        //}

        ////
        ////

        //private IEnumerable<string> GetLineFromWorksheet(ExcelObject.Worksheet excelWorksheet, int row)
        //{
        //    for (int col = 0; col < GetCountOfCols(excelWorksheet);col++)
        //    {
        //        yield return GetCellValue(excelWorksheet, row, col);
        //    }
        //}

        //
        //

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
            return excelWorksheet.Cells[row, col].Value.ToString();
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
        ///     Get count of actually filled rows in Worksheet
        /// </summary>
        /// <param name="excelWorksheet">
        ///     Worksheet to check
        /// </param>
        /// <returns></returns>
        public int GetCountOfRows(ExcelObject.Worksheet excelWorksheet)
        {
            return excelWorksheet.Cells[excelWorksheet.Rows.Count, 1].End(ExcelObject.XlDirection.xlUp).Row;
        }

        public int GetFirstRow(ExcelObject.Worksheet excelWorksheet)
        {
            int z = excelWorksheet.Cells[1, 1].End(ExcelObject.XlDirection.xlDown).Row;
            return z;
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
        ///     Get count of maximum actually filled columns in Worksheet
        /// </summary>
        /// <param name="excelWorksheet"></param>
        /// <returns></returns>
        public int GetCountOfCols(ExcelObject.Worksheet excelWorksheet)
        {
            return excelWorksheet.Cells[GetCountOfRows(excelWorksheet), excelWorksheet.Columns.Count].End(ExcelObject.XlDirection.xlToLeft).Column;
        }
    }
}
