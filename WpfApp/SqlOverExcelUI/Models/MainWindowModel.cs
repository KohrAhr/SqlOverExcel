using Lib.MVVM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlOverExcelUI.Models
{
    public class MainWindowModel : PropertyChangedNotification
    {
        public string ExcelFileName
        {
            get => GetValue(() => ExcelFileName);
            set => SetValue(() => ExcelFileName, value);
        }
    }
}
