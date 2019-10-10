using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace SqlOverExcelUI.Functions
{
    public class SearchValueConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (values.Length == 2)
            {
                string text = values[0].ToString().ToUpper();
                string searchText = values[1].ToString().ToUpper();

                if (!String.IsNullOrEmpty(searchText) && !String.IsNullOrEmpty(text))
                {
                    return text.Contains(searchText);
                }
            }

            return false;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

}
