using Hardcodet.Wpf.TaskbarNotification;
using Lib.Strings;
using Lib.System;
using Microsoft.Win32;
using SRPManagerV2.Core;
using SRPManagerV2.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace ExcelWorkbookSplitter
{
    public static class CoreFunctions
    {
        /// <summary>
        ///     Установить параметры командной строки
        /// </summary>
        /// <param name="commandLineParams"></param>
        public static bool ApplyCommandLineParameters(CommandLineParamsType commandLineParams)
        {
            // Специальный режим. Показать Help в консоле и выйти
            if (commandLineParams.ShowHelp)
            {
                Console.WriteLine("\n\n======================================");
                Console.WriteLine(StringsFunctions.ResourceString("resVersion"));
                Console.WriteLine("======================================\n");
                Console.WriteLine("Usage: SrpManager.exe [[[-master] | [-exit]] [[-enable] | [-disable]] [[-run_force] | [-stop_force]]] | [-?]");
                Console.WriteLine("\nOptions:");
                Console.WriteLine("\t-master    \tclose all other instance of application, and run new instance");
                Console.WriteLine("\t-exit      \tclose all other instance of application");
                Console.WriteLine("\t-enable    \tswitch SRP/AppLocker to the \"Whitelisting\" mode");
                Console.WriteLine("\t-disable   \tswitch SRP/AppLocker to the \"Blacklisting\" mode");
                Console.WriteLine("\t-force \tenable keeping of the selected mode while running");
                Console.WriteLine("\t-?         \tshow this help");
                ConsoleManager.FreeConsole();

                Application.Current.Shutdown();
                return false;
            }

            return true;
        }

        /// <summary>
        ///     Обработать параметры командной строки
        /// </summary>
        /// <param name="e">
        ///     Параметры командной строки
        /// </param>
        /// <returns>
        ///     Извлечённые значения
        /// </returns>
        public static CommandLineParamsType RecognizeCommandLineParams(StartupEventArgs e)
        {
            CommandLineParamsType result = new CommandLineParamsType();

            // proceed command line parameters
            foreach (string param in e.Args)
            {
                if (param.Contains("-?"))
                {
                    result.ShowHelp = true;
                }
                if (param.Contains("-master"))
                {
                    result.MasterInstance = true;
                }
                else
                if (param.Contains("-enable"))
                {
                    result.RequestedStatus = Status.sOff;
                }
                else
                if (param.Contains("-disable"))
                {
                    result.RequestedStatus = Status.sOn;
                }
                else
                if (param.Contains("-force"))
                {
                    result.ForceMode = true;
                }
                else
                if (param.Contains("-exit"))
                {
                    result.ExitRequest = true;
                }
                else
                if (param.Contains("-IgnoreMutex"/* + AppConsts.MUTEX_ID*/))
                {
                    result.IgnoreMutex = true;
                }
            }

            return result;
        }
    }
}
