using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Lib.Data
{
    public static class ClassFunctions
    {
        /// <summary>
        ///     Is text match in any fields in class
        ///     <para>If "filterText" is empty, return False (Not Match) -- mean that -- "No Entry" (or "empty string") always not equal to anything</para>
        ///     <para>This function could be used as alternative way for search in datagrid</para>
        /// </summary>
        /// <typeparam name="T">
        ///     Any type
        /// </typeparam>
        /// <param name="instance">
        ///     Instance of any class
        /// </param>
        /// <param name="text">
        ///     Text for search
        /// </param>
        /// <param name="caseSensative">
        ///     Default is False
        /// </param>
        /// <returns>
        ///     True if is match
        /// </returns>
        public static bool IsTextMatchInValues<T>(T instance, string text, bool caseSensative = false)
        {
            bool result = false;

            if (caseSensative)
            {
                text = text.ToUpper();
            }

            foreach (PropertyInfo propertyInfo in instance.GetType().GetProperties())
            {
                object value = propertyInfo.GetValue(instance, null);

                if (value == null)
                {
                    continue;
                }

                string valueAsString = value.ToString();

                if (caseSensative)
                {
                    valueAsString = valueAsString.ToUpper();
                }

                if (!String.IsNullOrEmpty(text))
                {
                    result = valueAsString.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;

                    // "If" for optimization
                    if (result)
                    {
                        break;
                    }
                }
            }

            return result;
        }
    }
}
