﻿using SqlOverExcelUI.Core;
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
            AppSettingsCore.Init();

            Current.ShutdownMode = ShutdownMode.OnMainWindowClose;

            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
        }

        private void AppExit(object sender, ExitEventArgs e)
        {
            // Thank you!
        }
    }
}