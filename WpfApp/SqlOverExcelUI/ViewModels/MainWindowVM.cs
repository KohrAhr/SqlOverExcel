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
using SqlOverExcelUI.Types;
using ExcelObject = Microsoft.Office.Interop.Excel;

namespace SqlOverExcelUI.ViewModels
{
    public class MainWindowVM
    {
        #region Commands definition
        public ICommand RunAnalyticsCommand { get; set; }
        public ICommand SelectFileCommand { get; set; }
        public ICommand RunSqlQueryCommand { get; set; }
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
            RunSqlQueryCommand = new RelayCommand(RunSqlQueryProc);
        }

        #region Commands implementation
        private void RunAnalyticsProc(object o)
        {
            using (ExcelCore excelIn = new ExcelCore(Model.ExcelFileName, "16.0"))
            {
                List<string> worksheets = excelIn.GetWorksheets();

                foreach (String name in worksheets)
                {
                    WorksheetItemType worksheetItem = new WorksheetItemType();

                    worksheetItem.WorksheetName = name;

                    ExcelObject.Worksheet worksheet = excelIn.GetWorksheet(name);
                    worksheetItem.RowCount = excelIn.GetCountOfRows(worksheet);
                    worksheetItem.ColCount = excelIn.GetCountOfCols(worksheet);

                    Model.WorksheetItems.Add(worksheetItem);
                }
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

        private void RunSqlQueryProc(object o)
        {
            // Execute query
        }
        #endregion Commands implementation
    }
}