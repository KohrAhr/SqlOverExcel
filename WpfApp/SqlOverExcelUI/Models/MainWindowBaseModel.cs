using Lib.MVVM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlOverExcelUI.Models
{
    public class MainWindowBaseModel : PropertyChangedNotification
    {
        /// <summary>
        ///     Excel file for proceed
        /// </summary>
        public string ExcelFileName
        {
            get => GetValue(() => ExcelFileName);
            set => SetValue(() => ExcelFileName, value);
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

        public string TextForSearch
        {
            get { return GetValue(() => TextForSearch); }
            set { SetValue(() => TextForSearch, value); }
        }
    }
}
