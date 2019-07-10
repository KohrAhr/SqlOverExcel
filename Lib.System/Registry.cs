using WpfSystem = System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.Win32;

namespace Lib.System
{
    public class RegistryFunctions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="valueName"></param>
        /// <param name="registry"></param>
        /// <returns>
        /// </returns>
        public string GetRegKeyValue(string path, string valueName, RegistryValueKind valueKind, RegistryKey registry = null)
        {
            string result = "";
            try
            {
                RegistryKey registryKey = null;

                if (registry == null)
                {
                    registryKey = Registry.LocalMachine;
                }

                using (RegistryKey key = registryKey.OpenSubKey(path))
                {
                    if (key != null)
                    {
                        RegistryValueKind registryValueKind = key.GetValueKind(valueName);

                        if (registryValueKind == valueKind)
                        {
                            WpfSystem.Object o = key.GetValue(valueName);
                            if (o != null)
                            {
                                if (o.GetType().BaseType == typeof(WpfSystem.Array))
                                {
                                    result = WpfSystem.String.Join(WpfSystem.Environment.NewLine, (string[])o);
                                }
                                else
                                {
                                    result = o.ToString();
                                }
                            }
                        }
                    }
                }
            }
            catch (WpfSystem.Exception)
            {
                result = null;
            }
            return result;
        }

        public string GetRegKeyValueObject(RegistryKey key, string valueName, RegistryValueKind valueKind)
        {
            string result = "";
            RegistryValueKind registryValueKind = key.GetValueKind(valueName);

            try
            {
                if (registryValueKind == valueKind)
                {
                    WpfSystem.Object o = key.GetValue(valueName);
                    if (o != null)
                    {
                        if (o.GetType().BaseType == typeof(WpfSystem.Array))
                        {
                            result = WpfSystem.String.Join(WpfSystem.Environment.NewLine, (string[])o);
                        }
                        else
                        {
                            result = o.ToString();
                        }
                    }
                }
            }
            catch (WpfSystem.NullReferenceException)
            {
                result = null;
            }
            catch (SecurityException)
            {
                result = null;
            }

            return result;
        }

        public long GetRegKeyValueObjectAsInt(RegistryKey key, string valueName, RegistryValueKind valueKind)
        {
            long result = 0;
            RegistryValueKind registryValueKind = key.GetValueKind(valueName);

            try
            {
                if (registryValueKind == valueKind)
                {
                    WpfSystem.Object o = key.GetValue(valueName);
                    if (o != null)
                    {
                        result = (long)o;
                    }
                }
            }
            catch (WpfSystem.NullReferenceException)
            {
                result = 0;
            }
            catch (SecurityException)
            {
                result = 0;
            }

            return result;
        }

        public bool CanSetRegKeyValue(string path, string valueName, RegistryKey registry = null)
        {
            bool result = true;

            try
            {
                RegistryKey registryKey = null;

                if (registry == null)
                {
                    registryKey = Registry.LocalMachine;
                }

                using (RegistryKey key = registryKey.OpenSubKey(path, true))
                {
                    result = key != null;
                }
            }
            catch (WpfSystem.NullReferenceException)
            {
                result = false;
            }
            catch (SecurityException)
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="path">
        ///     Must exist!
        /// </param>
        /// <param name="valueName"></param>
        /// <param name="value"></param>
        /// <param name="valueKind"></param>
        /// <param name="registry"></param>
        /// <returns></returns>
        public bool SetRegKeyValue(string path, string valueName, object value, RegistryValueKind valueKind, RegistryKey registry = null)
        {
            bool result = true;

            try
            {
                RegistryKey registryKey = null;

                if (registry == null)
                {
                    registryKey = Registry.LocalMachine;
                }

                using (RegistryKey key = registryKey.OpenSubKey(path, true))
                {
                    if (key != null)
                    {
                        //RegistryValueKind registryValueKind = key.GetValueKind(valueName);

                        //if (registryValueKind == valueKind)
                        //{
                        key.SetValue(valueName, value, valueKind);
//                        }
                    }
                }
            }
            catch (SecurityException)
            {
                result = false;
            }

            return result;
        }
    }
}
