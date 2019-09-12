using System;
using System.Collections.Generic;
using System.Data;
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
using System.Windows;

namespace SqlOverExcelUI.ViewModels
{
    public class MainWindowVM
    {
        private const string CONST_ACEOLEDBVERSION = "16.0";

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
            using (ExcelCore excelIn = new ExcelCore(Model.ExcelFileName, CONST_ACEOLEDBVERSION))
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
            openFileDialog.Filter = (string) Application.Current.FindResource("resFileTypes");
            openFileDialog.InitialDirectory = Environment.CurrentDirectory;
            if (openFileDialog.ShowDialog() == true)
            {
                Model.ExcelFileName = openFileDialog.FileName;
            }
        }

        private void RunSqlQueryProc(object o)
        {
            // Execute query

            DataTable queryResult = new DataTable();

            using (ExcelCore excelIn = new ExcelCore(Model.ExcelFileName, CONST_ACEOLEDBVERSION))
            {
                // Start query
                excelIn.RunSql(Model.SqlQuery, ref queryResult);

            }
        }
        #endregion Commands implementation
    }
}