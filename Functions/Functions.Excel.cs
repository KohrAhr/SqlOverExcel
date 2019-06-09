using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using ExcelObject = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookSplitter.Functions
{
    public class ExcelFile
    {
        public string FileName { get; set; }

        public ExcelObject.Application ExcelApp { set; get; } = null;
        public ExcelObject.Workbooks Books { set; get; } = null;
        public ExcelObject.Workbook Sheet { set; get; } = null;
        public ExcelObject.Worksheet Worksheet { set; get; } = null;
    }

    public class ExcelFunctions 
    {
        // ExcelFile, IDisposable
        // private ExcelFunctions excelFunctions = new ExcelFunctions();

        //public void Dispose()
        //{
        //    this.CloseFile(this);
        //}

        public ExcelFile OpenFile(string file)
        {
            ExcelFile result = new ExcelFile() { FileName = file };
            try
            {
                result.ExcelApp = new ExcelObject.Application();
                result.ExcelApp.Visible = false;
                result.Books = result.ExcelApp.Workbooks;
                result.Sheet = result.Books.Open(result.FileName);
            }
            catch
            {
                CloseFile(result);
                result = null;
            }

            return result;
        }

        public void CloseFile(ExcelFile excelFile)
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
        }

        public ExcelObject.Worksheet GetWorksheet(ExcelFile excelFile, String worksheetName)
        {
            return excelFile.Sheet.Worksheets[worksheetName];
        }

        public List<string> GetWorksheets(ExcelFile excelFile)
        {
            List<string> result = new List<string>();
            foreach (ExcelObject.Worksheet ws in excelFile.Sheet.Worksheets)
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

        //public GetExcelFile(String fileName)
        //{
        //    var file = new FileInfo(fileName);
        //    return new ExcelObject.ExcelPackage(file);
        //}

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
