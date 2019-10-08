using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Lib.MVVM;
using Microsoft.Win32;
using Lib.Excel;
using SqlOverExcelUI.Models;
using SqlOverExcelUI.Types;
using ExcelObject = Microsoft.Office.Interop.Excel;
using System.Windows;
using Lib.UI;
using Lib.Strings;
using System.Collections.ObjectModel;

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
        public ICommand UseTableNameCommand { get; set; }
        public ICommand AboutCommand { get; set; }
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
            RunAnalyticsCommand = new RelayCommand(RunAnalyticsProc, RunAnalyticsCommandEnabled);
            SelectFileCommand = new RelayCommand(SelectFileProc);
            RunSqlQueryCommand = new RelayCommand(RunSqlQueryProc, RunSqlQueryCommandEnabled);
            SaveQueryResultCommand = new RelayCommand(SaveQueryResultProc, SaveQueryResultCommandEnabled);
            UseTableNameCommand = new RelayCommand(UseTableNameProc, UseTableNameCommandEnabled);
            AboutCommand = new RelayCommand(AboutProc);
        }

        #region Commands implementation
        private void AboutProc(object o)
        {
            WindowsUI.RunWindowDialog(() =>
            {
                MessageBox.Show(
                    "SQL OVER EXCEL" + Environment.NewLine + "RIGA, LATVIA",
                    StringsFunctions.ResourceString("resInfo"),
                    MessageBoxButton.OK, MessageBoxImage.Information
                );
            }
            );
        }

        private void UseTableNameProc(object o)
        {
            if (o == null)
            {
                return;
            }

            WorksheetItemType selectedItem = (WorksheetItemType)((ObservableCollection<object>)o).FirstOrDefault();

            if (selectedItem == null)
            {
                return;
            }

            string s = selectedItem.WorksheetNameForQuery;

            Model.SqlQuery += s;
        }

        private bool UseTableNameCommandEnabled(Object o)
        {
            bool result = false;

            if (o != null)
            {
                WorksheetItemType selectedItem = (WorksheetItemType)((ObservableCollection<object>)o).FirstOrDefault();

                result = selectedItem != null;
            }

            return result;
        }

        /// <summary>
        ///     Save query result to new (prompted) Excel file
        /// </summary>
        /// <param name="o"></param>
        private void SaveQueryResultProc(object o)
        {
            // Ask for a file name
            string fileName = "";

            WindowsUI.RunWindowDialog(() =>
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = StringsFunctions.ResourceString("resFileTypes");
                saveFileDialog.InitialDirectory = Environment.CurrentDirectory;
                if (saveFileDialog.ShowDialog() == true)
                {
                    fileName = saveFileDialog.FileName;
                }
            }
            );

            if (String.IsNullOrEmpty(fileName))
            {
                return;
            }

            // Save
            using (new WaitCursor())
            {
                new ExcelCore().SaveResultToExcelFile(fileName, Model.QueryResult);
            }

            WindowsUI.RunWindowDialog(() =>
            {
                MessageBox.Show(
                    String.Format(StringsFunctions.ResourceString("resResultSaved"), fileName),
                    StringsFunctions.ResourceString("resInfo"),
                    MessageBoxButton.OK, MessageBoxImage.Information
                );
            }
            );
        }

        private bool SaveQueryResultCommandEnabled(Object o)
        {
            return Model.QueryResult.Rows.Count > 0;
        }

        private void RunAnalyticsProc(object o)
        {
            try
            {
                using (new WaitCursor())
                {
                    using (ExcelCore excelIn = new ExcelCore(Model.ExcelFileName, CONST_ACEOLEDBVERSION, Model.HDR))
                    {
                        List<string> worksheets = excelIn.GetWorksheets();

                        Model.WorksheetItems.Clear();

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

        private bool RunAnalyticsCommandEnabled(object o)
        {
            return !String.IsNullOrEmpty(Model.ExcelFileName);
        }

        private void SelectFileProc(object o)
        {
            WindowsUI.RunWindowDialog(() =>
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = StringsFunctions.ResourceString("resFileTypes");
                openFileDialog.InitialDirectory = Environment.CurrentDirectory;
                if (openFileDialog.ShowDialog() == true)
                {
                    Model.ExcelFileName = openFileDialog.FileName;
                }
            }
            );
        }

        private void RunSqlQueryProc(object o)
        {
            try
            {
                using (new WaitCursor())
                {
                    Task runQuery = new Task(() =>
                    {
                        using (ExcelCore excelIn = new ExcelCore(Model.ExcelFileName, CONST_ACEOLEDBVERSION, Model.HDR))
                        {
                            DataTable queryResult = new DataTable();

                            // Run query
                            excelIn.RunSql(Model.SqlQuery, ref queryResult);

                            // Populate result
                            Model.QueryResult = queryResult;
                        }
                    });

                    runQuery.Start();
                    runQuery.Wait();
                }

                WindowsUI.RunWindowDialog(() =>
                {
                    MessageBox.Show(
                        StringsFunctions.ResourceString("resQueryCompleted"),
                        StringsFunctions.ResourceString("resInfo"),
                        MessageBoxButton.OK, MessageBoxImage.Information
                    );
                }
                );
            }
            catch (Exception ex)
            {
                WindowsUI.RunWindowDialog(() =>
                {
                    MessageBox.Show(
                        StringsFunctions.ResourceString("resErrorDuringExecution") + 
                            Environment.NewLine + Environment.NewLine + ex.InnerException.Message,
                        StringsFunctions.ResourceString("resError"),
                        MessageBoxButton.OK, MessageBoxImage.Hand
                    );
                }
                );
            }
        }
        private bool RunSqlQueryCommandEnabled(object o)
        {
            return !String.IsNullOrEmpty(Model.SqlQuery);
        }
        #endregion Commands implementation
    }
}