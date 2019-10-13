using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace SqlOverExcelUI.Core
{
    public static class AppSettingsCore
    {
        /// <summary>
        ///     Default value
        /// </summary>
        private const string CONST_ACEOLEDBVERSION = "16.0";

        /// <summary>
        ///     Project name -- folder name with Config file
        /// </summary>
        private const string CONST_PRODUCT_NAME = "SqlOverExcel";

        /// <summary>
        ///     Configuration file
        /// </summary>
        private const string CONST_CONFIG_FILE_NAME = "SqlOverExcel.config";

        /// <summary>
        ///     Full path and file name of configuration file
        /// </summary>
        public static string SettingsFile = Path.Combine(
            Environment.ExpandEnvironmentVariables("%APPDATA%"),
            CONST_PRODUCT_NAME,
            CONST_CONFIG_FILE_NAME
        );

        /// <summary>
        ///     Constructor
        /// </summary>
        public static void Init()
        {
			if (File.Exists(SettingsFile))
            {
                LoadSettings(SettingsFile);
            }
        }

		public static void LoadSettings(string FileName)
        {
            ExeConfigurationFileMap map = new ExeConfigurationFileMap();
            map.ExeConfigFilename = SettingsFile;

            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);

            string value = "";

            value = config.AppSettings.Settings["AceVer"]?.Value.ToString();
            if (String.IsNullOrEmpty(value))
            {
                value = CONST_ACEOLEDBVERSION;
            }
            AppDataCore.AceVersion = value;
        }
    }
}
