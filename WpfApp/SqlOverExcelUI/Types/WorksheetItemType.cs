using Lib.MVVM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlOverExcelUI.Types
{
    /// <summary>
    ///     Base "Worksheet Item" class
    /// </summary>
    public class WorksheetItemType : PropertyChangedNotification, ICloneable
    {
        public string WorksheetName
        {
            get => GetValue(() => WorksheetName);
            set => SetValue(() => WorksheetName, value);
        }

        public int RowCount
        {
            get => GetValue(() => RowCount);
            set => SetValue(() => RowCount, value);
        }

        public int ColCount
        {
            get => GetValue(() => ColCount);
            set => SetValue(() => ColCount, value);
        }

        /// <summary>
        ///  
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            return new WorksheetItemType()
            {
                WorksheetName = this.WorksheetName,
                RowCount = this.RowCount,
                ColCount = this.ColCount
            };
        }
    }
}
