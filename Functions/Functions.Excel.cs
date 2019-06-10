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
        ///     Class with information about current Excel file/object
        /// </summary>
        protected internal ExcelFile excelFile = null;

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

        public void OpenFile(string file)
        {
            excelFile = new ExcelFile()
            {
                FileName = file
            };

            try
            {
                excelFile.ExcelApp = new ExcelObject.Application();
                excelFile.ExcelApp.Visible = false;
                excelFile.Books = excelFile.ExcelApp.Workbooks;
                excelFile.Sheet = excelFile.Books.Open(file);
            }
            catch
            {
                CloseFile();
                excelFile = null;
            }
        }

        public void CloseFile()
        {
            if (excelFile?.Worksheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFile.Worksheet);
            }
            excelFile?.Sheet?.Close(false);
            excelFile?.Books?.Close();
            if (excelFile?.Sheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFile.Sheet);
            }
            if (excelFile?.Books != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFile.Books);
            }
            excelFile?.ExcelApp?.Quit();
            if (excelFile?.ExcelApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFile?.ExcelApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public ExcelObject.Worksheet GetWorksheet(ExcelFile excelFile, String worksheetName)
        {
            return excelFile.Sheet.Worksheets[worksheetName];
        }

        public ExcelObject.Worksheet GetWorksheet(ExcelFile excelFile, int worksheetId)
        {
            return excelFile.Sheet.Worksheets[worksheetId];
        }

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

        public String GetCellValue(ExcelObject.Worksheet excelWorksheet, int row, int col)
        {
            return excelWorksheet.Cells[row, col].Value.ToString();
        }

        //

        public int GetCountOfRows(ExcelObject.Worksheet excelWorksheet)
        {
            return excelWorksheet.Rows.Count;
        }

        public int GetCountOfCols(ExcelObject.Worksheet excelWorksheet)
        {
            return excelWorksheet.Columns.Count;
        }
    }
}
