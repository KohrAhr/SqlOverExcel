using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lib.Strings
{
    public static class StringsFunctions
    {
        /// <summary>
        ///     Convert string with Hex value to array of bytes
        ///     <para>https://stackoverflow.com/questions/321370/how-can-i-convert-a-hex-string-to-a-byte-array</para>
        /// </summary>
        /// <param name="hex">Hex values as string</param>
        /// <returns>Byte array</returns>
        public static byte[] StringToByteArray(string hex)
        {
            return Enumerable.Range(0, hex.Length / 2).Select(x => Convert.ToByte(hex.Substring(x * 2, 2), 16)).ToArray();
        }

        /// <summary>
        ///     Find resource by name and return value
        /// </summary>
        /// <param name="name">
        ///     Resource name
        /// </param>
        /// <returns>
        ///     <para>If resource not found, will be returned name</para>
        /// </returns>
        public static string ResourceString(string name)
        {
            object o = System.Windows.Application.Current?.TryFindResource(name);

            return o == null ? name : o.ToString();
        }

        /// <summary>
        ///     Convert UTF8 string to array of bytes
        /// </summary>
        /// <param name="value">
        ///     UTF8 String for encoding
        /// </param>
        /// <returns>
        ///     Encoded UTF8 string as array of bytes
        /// </returns>
        public static byte[] StringToUtf8Bytes(string value)
        {
            byte[] result = null;
            if (!String.IsNullOrEmpty(value))
            {
                result = Encoding.UTF8.GetBytes(value);
            }
            return result;
        }

        /// <summary>
        ///     Validation that string contain only digits and nothing else
        /// </summary>
        /// <param name="value">
        ///     String for validation
        /// </param>
        /// <returns></returns>
        public static bool StringContainOnlyDigits(string value)
        {
            bool result = false;
            if (!String.IsNullOrEmpty(value))
            {
                result = value.All(Char.IsDigit);
            }
            return result;
        }

        /// <summary>
        ///     Just special formatting
        ///     Using only for debug purpose! Because I cannot proof that encryption & description works as it should be.
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static string BytesAsHexString(byte[] bytes, string splitBy = "")
        {
            StringBuilder result = new StringBuilder((bytes.Length + splitBy.Length) * 2);
            foreach (byte b in bytes)
            {
                if (!String.IsNullOrEmpty(result.ToString()))
                {
                    result.AppendFormat("{0}", splitBy);
                }
                result.AppendFormat("{0:X2}", b);
            }
            return result.ToString().Trim();
        }

        /// <summary>
        ///     Just special formatting
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static string BytesAsString(byte[] bytes, string splitBy = "")
        {
            StringBuilder result = new StringBuilder((bytes.Length + splitBy.Length) * 4);
            foreach (byte b in bytes)
            {
                if (!String.IsNullOrEmpty(result.ToString()))
                {
                    result.AppendFormat("{0}", splitBy);
                }
                result.AppendFormat("{0}", b);
            }
            return result.ToString().Trim();
        }
    }
}
