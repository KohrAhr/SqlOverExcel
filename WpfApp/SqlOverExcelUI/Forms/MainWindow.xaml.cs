using SqlOverExcelUI.Models;
using SqlOverExcelUI.Types;
using SqlOverExcelUI.ViewModels;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SqlOverExcelUI.Functions;
using System.Xml;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Folding;

namespace SqlOverExcelUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            DataContext = new MainWindowVM();

            string resName = "sql.xshd";

            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resName))
            {
                using (XmlTextReader reader = new XmlTextReader(resName))
                {
                    txtCode.SyntaxHighlighting =
                        ICSharpCode.AvalonEdit.Highlighting.Xshd.HighlightingLoader.Load(reader, HighlightingManager.Instance);
                }
            }

            txtCode.ShowLineNumbers = true;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBoxName = (TextBox)sender;
            string filterText = textBoxName.Text.ToUpper();

            FilterGridView(dgResult.ItemsSource, filterText);
        }
        private void FilterGridView(IEnumerable itemsControls, string filterText)
        {
            ICollectionView cv = CollectionViewSource.GetDefaultView(itemsControls);

            //cv.Filter = o => {
            //    return true;
            //};
        }

    }
}
