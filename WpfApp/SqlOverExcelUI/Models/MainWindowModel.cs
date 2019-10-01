using Lib.MVVM;
using SqlOverExcelUI.Types;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlOverExcelUI.Models
{
    /// <summary>
    ///     Model for Main Window
    /// </summary>
    public class MainWindowModel : PropertyChangedNotification
    {
        /// <summary>
        ///     Constructor
        /// </summary>
        public MainWindowModel()
        {
            WorksheetItems = new WorksheetItemsType();
            QueryResult = new DataTable();
            HDR = true;
        }

        /// <summary>
        ///     Destructor
        /// </summary>
        ~MainWindowModel()
        {
            GC.Collect();
        }

        /// <summary>
        ///     Excel file for proceed
        /// </summary>
        public string ExcelFileName
        {
            get => GetValue(() => ExcelFileName);
            set => SetValue(() => ExcelFileName, value);
        }

        /// <summary>
        ///     Data table contain SQL query execution result
        /// </summary>
        public DataTable QueryResult
        {
            get { return GetValue(() => QueryResult); }
            set { SetValue(() => QueryResult, value); }
        }

        /// <summary>
        ///     Information about all worksheets
        /// </summary>
        public WorksheetItemsType WorksheetItems
        {
            get { return GetValue(() => WorksheetItems); }
            set { SetValue(() => WorksheetItems, value); }
        }

        /// <summary>
        ///     SQL query to run
        /// </summary>
        public string SqlQuery
        {
            get { return GetValue(() => SqlQuery); }
            set { SetValue(() => SqlQuery, value); }
        }

        public bool HDR
        {
            get { return GetValue(() => HDR); }
            set { SetValue(() => HDR, value); }
        }

    }
}
