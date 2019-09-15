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
using Lib.UI;
using Lib.Strings;

namespace SqlOverExcelUI.ViewModels
{
    public class MainWindowVM
    {
        private const string CONST_ACEOLEDBVERSION = "16.0";

        #region Commands definition
        public ICommand RunAnalyticsCommand { get; set; }
        public ICommand SelectFileCommand { get; set; }
        public ICommand RunSqlQueryCommand { get; set; }
        public ICommand SaveQueryResultCommand { get; set; }
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
            SaveQueryResultCommand = new RelayCommand(SaveQueryResultProc);
        }

        #region Commands implementation
        /// <summary>
        ///     Save query result to new (prompted) Excel file
        /// </summary>
        /// <param name="o"></param>
        private void SaveQueryResultProc(object o)
        {
            // Ask for a file name

            // DataTable to CSV ?

            // DataTable to Excel ?

        }

        private void RunAnalyticsProc(object o)
        {
            try
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
            catch (Exception ex)
            {
                WindowsUI.RunWindowDialog(() =>
                {
                    MessageBox.Show(
                        String.Format(StringsFunctions.ResourceString("resErrorDuringOpening"), Model.ExcelFileName) +
                            Environment.NewLine + Environment.NewLine + ex.Message,
                        StringsFunctions.ResourceString("resError"),
                        MessageBoxButton.OK, MessageBoxImage.Hand
                    );
                }
                );
            }
        }

        private void SelectFileProc(object o)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = StringsFunctions.ResourceString("resFileTypes");
            openFileDialog.InitialDirectory = Environment.CurrentDirectory;
            if (openFileDialog.ShowDialog() == true)
            {
                Model.ExcelFileName = openFileDialog.FileName;
            }
        }

        private void RunSqlQueryProc(object o)
        {
            try
            {
                using (ExcelCore excelIn = new ExcelCore(Model.ExcelFileName, CONST_ACEOLEDBVERSION))
                {
                    DataTable queryResult = new DataTable();

                    // Run query
                    excelIn.RunSql(Model.SqlQuery, ref queryResult);

                    // Populate result
                    Model.QueryResult = queryResult;
                }
            }
            catch (Exception ex)
            {
                WindowsUI.RunWindowDialog(() =>
                {
                    MessageBox.Show(
                        StringsFunctions.ResourceString("resErrorDuringExecution") + 
                            Environment.NewLine + Environment.NewLine + ex.Message,
                        StringsFunctions.ResourceString("resError"),
                        MessageBoxButton.OK, MessageBoxImage.Hand
                    );
                }
                );
            }
        }
        #endregion Commands implementation
    }
}