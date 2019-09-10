using Lib.MVVM;
using SqlOverExcelUI.Types;
using System;
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

        public WorksheetItemsType WorksheetItems
        {
            get { return GetValue(() => WorksheetItems); }
            set { SetValue(() => WorksheetItems, value); }
        }
    }
}
