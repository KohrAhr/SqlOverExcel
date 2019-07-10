using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Lib.System
{
    public static class CommandLineFunctions
    {
        public static Dictionary<string, string> GetCommandLineParameters(StartupEventArgs startupEventArgs)
        {
            return GetCommandLineParameters(startupEventArgs.Args);
        }

        public static Dictionary<string, string> GetCommandLineParameters(string[] request)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            if (request != null && request.Length > 0)
            {
                foreach (string currentKeyValuePair in request)
                {
                    int index = currentKeyValuePair.IndexOf("=");

                    if (index < 0)
                    {
                        continue;
                    }

                    string keyName = currentKeyValuePair.Substring(0, index).ToLower().Trim();
                    string keyValue = currentKeyValuePair.Substring(index + 1, currentKeyValuePair.Length - 1 - index);

                    // No name? Ignore. Value could be empty, but not Key
                    if (String.IsNullOrEmpty(keyName))
                    {
                        continue;
                    }

                    result.Add(keyName, keyValue);
                }
            }

            return result;
        }

    }
}
