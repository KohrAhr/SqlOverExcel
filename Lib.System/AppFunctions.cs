using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace Lib.System
{
    public static class AppFunctions
    {
        public static string GetAppBuild()
        {
            return Assembly.GetEntryAssembly().GetName().Version.ToString();
        }
    }
}
