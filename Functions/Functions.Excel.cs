using System;

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
        private ExcelFile excelFile = null;

        /// <summary>
        ///     Class with Excel functions/API
        /// </summary>
        private ExcelFunctions excelFunctions = null;

        /// <summary>
        ///     Default constructor
        /// </summary>
        public ExcelCore()
        {
            excelFunctions = new ExcelFunctions();
        }

        /// <summary>
        ///     Constructor with file name
        /// </summary>
        /// <param name="FileName">
        ///     Excel file name
        /// </param>
        public ExcelCore(String FileName) : this()
        {
            excelFile = excelFunctions.OpenFile(FileName);
        }

        /// <summary>
        ///     Destructor
        /// </summary>
        public void Dispose()
        {
            excelFunctions?.CloseFile(excelFile);
        }
    }
}
