using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace SqlOverExcelUI.Functions
{
    /// <summary>
    ///     Highlight. Modern way
    /// </summary>
    public static class DataGridTextSearch
    {
        // Using a DependencyProperty as the backing store for SearchValue.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SearchValueProperty =
            DependencyProperty.RegisterAttached("SearchValue", typeof(string), typeof(DataGridTextSearch),
                new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.Inherits));

        public static string GetSearchValue(DependencyObject obj)
        {
            return (string)obj.GetValue(SearchValueProperty);
        }

        public static void SetSearchValue(DependencyObject obj, string value)
        {
            obj.SetValue(SearchValueProperty, value);
        }

        // Using a DependencyProperty as the backing store for IsTextMatch.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsTextMatchProperty =
            DependencyProperty.RegisterAttached("IsTextMatch", typeof(bool), typeof(DataGridTextSearch), new UIPropertyMetadata(false));

        public static bool GetIsTextMatch(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsTextMatchProperty);
        }

        public static void SetIsTextMatch(DependencyObject obj, bool value)
        {
            obj.SetValue(IsTextMatchProperty, value);
        }
    }

    /// <summary>
    ///     Not smart way
    ///     Delete
    /// </summary>
    public static class DataGridFunctions
    {
        public static void ApplyFilterOnDataGrid(string query, DataGrid dataGrid)
        {
            ICollectionView cv = CollectionViewSource.GetDefaultView(dataGrid.ItemsSource);
            if (String.IsNullOrEmpty(query))
            {
                cv.Filter = null;
            }
            else
            {
                cv.Filter = x =>
                {
                    bool match = false;
                    foreach (PropertyInfo propertyInfo in x.GetType().GetProperties())
                    {
                        object value = propertyInfo.GetValue(x, null);

                        if (value == null)
                        {
                            continue;
                        }

                        string valueAsString = value.ToString();

                        match = valueAsString.ToUpper().Contains(query.ToUpper());

                        // "If" for optimization
                        if (match)
                        {
                            break;
                        }
                    }

                    return match;
                };
            }
        }
    }
}
