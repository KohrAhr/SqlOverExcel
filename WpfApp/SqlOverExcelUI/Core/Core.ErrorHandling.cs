using Lib.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SqlOverExcelUI.Core
{
    public static class ErrorHandlingCore
    {
        /// <summary>
        /// 
        /// </summary>
        public static void Init()
        {
            // Set Error catching & handling
            AppDomain.CurrentDomain.UnhandledException += (s, eventArgs) =>
                {
                    // TODO: All data to File by NLog. 

                    // Final message to user
                    // In some cases (broken styles) message box cannot be displayed

                    WindowsUI.RunWindowDialog(() =>
                    {
                        MessageBox.Show(
                            eventArgs.ExceptionObject.ToString(),
                            "Error Stack Trace",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error
                        );
                    }
                    );

                    // Possible
                    //Application.Current.Shutdown();
                };
        }
    }
}
