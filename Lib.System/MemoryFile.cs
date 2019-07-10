using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Security.AccessControl;

namespace Lib.System
{
    public static class MMFFunctions
    {
        /// <summary>
        ///     Проверить что MemoryMappedFile существует
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool Exist(string name)
        {
            bool result = false;

            try
            {
                using (MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite, HandleInheritability.Inheritable))
                {

                }
                result = true;
            }
            catch (FileNotFoundException)
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        ///     Создать MemoryMappedFile
        /// </summary>
        /// <param name="name"></param>
        /// <param name="lenght"></param>
        /// <returns></returns>
        public static MemoryMappedFile Create(string name, int lenght)
        {
            try
            {
                return MemoryMappedFile.CreateNew(name, lenght, MemoryMappedFileAccess.ReadWriteExecute, MemoryMappedFileOptions.None, GetMemoryMappedFileSecurity(), HandleInheritability.Inheritable);
            }
            catch (UnauthorizedAccessException)
            {
                throw new UnauthorizedAccessException();
            }
        }

        /// <summary>
        ///     Открыть существующий MemoryMappedFile
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static MemoryMappedFile Open(string name)
        {
            return MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite);
        }

        public static MemoryMappedFileSecurity GetMemoryMappedFileSecurity()
        {
            MemoryMappedFileSecurity result = new MemoryMappedFileSecurity();

            result.AddAccessRule(new AccessRule<MemoryMappedFileRights>(
                    new SecurityIdentifier(WellKnownSidType.WorldSid, null),
                    MemoryMappedFileRights.FullControl,
                    AccessControlType.Allow
                )
            );

            return result;
        }
    }
}
