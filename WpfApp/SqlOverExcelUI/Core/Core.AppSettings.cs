using SqlOverExcelUI.Models;
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
            AppSettingsModel result = null;

            if (File.Exists(SettingsFile))
            {
                result = LoadSettings(SettingsFile);
            }

            AppDataCore.Settings.AceVersion = result == null ? CONST_ACEOLEDBVERSION : result.AceVersion;
        }

        public static AppSettingsModel LoadSettings(string FileName)
        {
            AppSettingsModel result = new AppSettingsModel();

            ExeConfigurationFileMap map = new ExeConfigurationFileMap();
            map.ExeConfigFilename = SettingsFile;

            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);

            result.AceVersion = config.AppSettings.Settings["AceVer"]?.Value.ToString();

            return result;
        }
    }
}
