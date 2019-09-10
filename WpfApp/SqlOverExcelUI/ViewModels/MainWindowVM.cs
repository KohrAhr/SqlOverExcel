using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Lib.MVVM;
using SqlOverExcelUI.Models;

namespace SqlOverExcelUI.ViewModels
{
    public class MainWindowVM
    {
        #region Commands definition
        public ICommand RunCommand { get; set; }
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
            Model.ExcelFileName = "<test>";
        }

        private void InitCommands()
        {
            RunCommand = new RelayCommand(RunCommandProc);
            SelectFileCommand = new RelayCommand(SelectFileProc);
        }

        #region Commands implementation
        private void RunCommandProc(object o)
        {

        }

        private void SelectFileProc(object o)
        {

        }
        #endregion Commands implementation
    }
}