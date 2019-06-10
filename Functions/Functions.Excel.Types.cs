using ExcelObject = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookSplitter.Functions
{
    /// <summary>
    ///     Structure contain different kind of Excel objects 
    ///     and 
    ///     common information about Excel file (file name)
    /// </summary>
    public class ExcelFile
    {
        /// <summary>
        ///     Excel file name
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        ///     Excel application
        /// </summary>
        protected internal ExcelObject.Application ExcelApp { set; get; } = null;

        /// <summary>
        /// 
        /// </summary>
        protected internal ExcelObject.Workbooks Books { set; get; } = null;

        /// <summary>
        /// 
        /// </summary>
        protected internal ExcelObject.Workbook Sheet { set; get; } = null;

        /// <summary>
        /// 
        /// </summary>
        protected internal ExcelObject.Worksheet Worksheet { set; get; } = null;
    }
}
