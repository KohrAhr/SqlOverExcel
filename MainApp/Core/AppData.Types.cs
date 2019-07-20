using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWorkbookSplitter.Core
{
    /// <summary>
    ///     Class with major settings
    /// </summary>
    public class AppData
    {
        /// <summary>
        ///     File to proceed (incomming file)
        /// </summary>
        public string inFile = "";

        /// <summary>
        ///     Output file
        /// </summary>
        public string outFile = "";

        /// <summary>
        ///     SQL query to run
        /// </summary>
        public string query = "";

        /// <summary>
        ///     ACE.OLEDB version
        /// </summary>
        public string acever = "16.0";

        /// <summary>
        /// 
        /// </summary>
        public bool showHelp
        {
            get
            {
                return String.IsNullOrEmpty(inFile);
            }
        }

        /// <summary>
        ///     Display common information about worksheets in Excel file
        /// </summary>
        public bool showInfo
        {
            get
            {
                return String.IsNullOrEmpty(query);
            }
        }

        /// <summary>
        ///     Save Query result to file
        ///     <para>False -- display information in console</para>
        /// </summary>
        public bool resultToFile
        {
            get
            {
                return !String.IsNullOrEmpty(outFile);
            }
        }
    }
}
