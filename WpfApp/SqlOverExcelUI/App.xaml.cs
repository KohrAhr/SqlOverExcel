using SqlOverExcelUI.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace SqlOverExcelUI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        void AppStartup(Object sender, StartupEventArgs e)
        {
            ErrorHandlingCore.Init();

            Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;
        }
    }
}