using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Lib.MVVM;
using Microsoft.Win32;
using SqlOverExcelUI.Functions;
using SqlOverExcelUI.Models;

namespace SqlOverExcelUI.ViewModels
{
    public class MainWindowVM
    {
        #region Commands definition
        public ICommand RunAnalyticsCommand { get; set; }
        public ICommand SelectFileCommand { get; set; }
        #endregion Commands definition

        public MainWindowModel Model
        {
            get; set;
        }

        public MainWindowVM()
        {
            InitData();

            InitCommands();
        }

        private void InitData()
        {
            Model = new MainWindowModel();
        }

        private void InitCommands()
        {
            RunAnalyticsCommand = new RelayCommand(RunAnalyticsProc);
            SelectFileCommand = new RelayCommand(SelectFileProc);
        }

        #region Commands implementation
        private void RunAnalyticsProc(object o)
        {
            using (ExcelCore excelIn = new ExcelCore(Model.ExcelFileName, "16.0"))
            {
                List<string> worksheets = excelIn.GetWorksheets();

                //foreach (String name in worksheets)
                //{
                //    Console.WriteLine("\t\"{0}\"", name);

                //    ExcelObject.Worksheet worksheet = excel.GetWorksheet(name);

                //    Console.WriteLine("\tLast row with data: {0}; Last column with data: {1}\n",
                //        excel.GetCountOfRows(worksheet),
                //        excel.GetCountOfCols(worksheet)
                //    );
                //}
            }
        }

        private void SelectFileProc(object o)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel Binary Workbook|(*.xlsb)|Excel 97-2003 Workbook|(*.xls)|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.CurrentDirectory;
            if (openFileDialog.ShowDialog() == true)
            {
                Model.ExcelFileName = openFileDialog.FileName;
            }
        }
        #endregion Commands implementation
    }
}