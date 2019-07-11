﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWorkbookSplitter.Core
{
    /// <summary>
    ///     Static class with major settings
    /// </summary>
    public static class AppData
    {
        /// <summary>
        ///     File to proceed (incomming file)
        /// </summary>
        public static string inFile = "";

        /// <summary>
        ///     Output file
        /// </summary>
        public static string outFile = "";

        /// <summary>
        ///     SQL query to run
        /// </summary>
        public static string query = "";

        /// <summary>
        ///     Display common information about worksheets in Excel file
        /// </summary>
        public static bool showInfo = false;

        /// <summary>
        ///     Save Query result to file
        ///     <para>False -- display information in console</para>
        /// </summary>
        public static bool resultToFile
        {
            get
            {
                return !String.IsNullOrEmpty(outFile);
            }
        }
    }
}
