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
    public class MainWindowModel : PropertyChangedNotification
    {
        public MainWindowModel()
        {
            WorksheetItems = new WorksheetItemsType();
            QueryResult = new DataTable();
        }

        ~MainWindowModel()
        {
            GC.Collect();
        }

        public string ExcelFileName
        {
            get => GetValue(() => ExcelFileName);
            set => SetValue(() => ExcelFileName, value);
        }

        public DataTable QueryResult
        {
            get { return GetValue(() => QueryResult); }
            set { SetValue(() => QueryResult, value); }
        }

        public WorksheetItemsType WorksheetItems
        {
            get { return GetValue(() => WorksheetItems); }
            set { SetValue(() => WorksheetItems, value); }
        }
        public string SqlQuery
        {
            get { return GetValue(() => SqlQuery); }
            set { SetValue(() => SqlQuery, value); }
        }
    }
}
