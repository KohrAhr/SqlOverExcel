using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lib.System
{
    public class ProcessesFunctions    
    {
        private Process _RunAs(string file, string args, string verb = "")
        {
            // Setting up start info of the new process of the same application
            ProcessStartInfo processStartInfo = new ProcessStartInfo(file);

            // Using operating shell and setting the ProcessStartInfo.Verb to “runas” will let it run as admin
            processStartInfo.UseShellExecute = true;
            processStartInfo.Verb = verb;
            processStartInfo.Arguments = args;

            // Start the application as new process
            Process process = null;
            try
            {
                process = Process.Start(processStartInfo);
            }
            catch
            {
                process = null;
            }

            return process;
        }

        public Process RunAs(string file, string args = "")
        {
            return _RunAs(file, args, "runas");
        }

        public Process Run(string file, string args = "")
        {
            return _RunAs(file, args);
        }
    }
}
