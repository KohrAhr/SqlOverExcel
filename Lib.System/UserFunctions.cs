using System;
using System.Security;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Security.Principal;

namespace Lib.System
{
    public static class UserFunctions
    {
        public static string GetDomainName()
        {
            string result = "";

            try
            {
                Domain domain = Domain.GetComputerDomain();
                result = domain.Name;
            }
            catch (ActiveDirectoryObjectNotFoundException)
            {
                result = "";
            }

            return result;
        }

        public static string GetPcName()
        {
            string result = "";

            try
            {
                result = Environment.MachineName;
            }
            catch (InvalidOperationException)
            {
                result = "";
            }

            return result;
        }

        public static string GetUsername()
        {
            return Environment.UserName;
        }

        public static string GetDomainOrPcName()
        {
            return InDomain() ? GetDomainName() : GetPcName();
        }

        public static string GetFQDN()
        {
            return 
                String.Format(
                    @"{0}\{1}",
                    GetDomainOrPcName(), 
                    GetUsername()
                );
        }

        /// <summary>
        ///     Validate username and password combination    
        ///     <para>Following Windows Services must be up</para>
        ///     <para>LanmanServer; TCP/IP NetBIOS Helper</para>
        /// </summary>
        /// <param name="userName">
        ///     Fully formatted UserName.
        ///     In AD: Domain + Username
        ///     In Workgroup: Username or Local computer name + Username
        /// </param>
        /// <param name="securePassword"></param>
        /// <returns></returns>
        public static bool ValidateUsernameAndPassword(string userName, SecureString securePassword)
        {
            bool result = false;

            ContextType contextType = ContextType.Machine;

            if (InDomain())
            {
                contextType = ContextType.Domain;
            }

            try
            {
                using (PrincipalContext principalContext = new PrincipalContext(contextType))
                {
                    result = principalContext.ValidateCredentials(
                        userName, 
                        new NetworkCredential(string.Empty, securePassword).Password
                    );
                }
            }
            catch (PrincipalOperationException)
            {
                // Account disabled? Considering as Login failed
                result = false;
            }
            catch (Exception)
            {
                throw;
            }

            return result;
        }


        /// <summary>
        ///     Validate: computer connected to domain?   
        /// </summary>
        /// <returns>
        ///     True -- computer is in domain
        ///     <para>False -- computer not in domain</para>
        /// </returns>
        public static bool InDomain()
        {
            bool result = true;

            try
            {
                Domain domain = Domain.GetComputerDomain();
            }
            catch (ActiveDirectoryObjectNotFoundException)
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        ///     
        /// </summary>
        /// <returns></returns>
        //public static bool CurrentUserHasAdminRights()
        //{
        //    return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
        //}

        /// <summary>
        /// Determines whether the specified user is an administrator.
        /// </summary>
        /// <param name="username">The user name.</param>
        /// <returns>
        ///   <c>true</c> if the specified user is an administrator; otherwise, <c>false</c>.
        /// </returns>
        /// <seealso href="https://ayende.com/blog/158401/are-you-an-administrator"/>
        public static bool IsCurrentUserAdmin(bool checkCurrentRole = true)
        {
            bool isElevated = false;

            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                if (checkCurrentRole)
                {
                    // Even if the user is defined in the Admin group, UAC defines 2 roles: one user and one admin. 
                    // IsInRole consider the current default role as user, thus will return false!
                    // Will consider the admin role only if the app is explicitly run as admin!
                    WindowsPrincipal principal = new WindowsPrincipal(identity);
                    isElevated = principal.IsInRole(WindowsBuiltInRole.Administrator);
                }
                else
                {
                    // read all roles for the current identity name, asking ActiveDirectory
                    isElevated = IsAdministratorNoCache(identity.Name);
                }
            }

            return isElevated;
        }

        /// <summary>
        /// Determines whether the specified user is an administrator.
        /// </summary>
        /// <param name="username">The user name.</param>
        /// <returns>
        ///   <c>true</c> if the specified user is an administrator; otherwise, <c>false</c>.
        /// </returns>
        /// <seealso href="https://ayende.com/blog/158401/are-you-an-administrator"/>
        private static bool IsAdministratorNoCache(string username)
        {
            PrincipalContext ctx;
            try
            {
                Domain.GetComputerDomain();
                try
                {
                    ctx = new PrincipalContext(ContextType.Domain);
                }
                catch (PrincipalServerDownException)
                {
                    // can't access domain, check local machine instead 
                    ctx = new PrincipalContext(ContextType.Machine);
                }
            }
            catch (ActiveDirectoryObjectNotFoundException)
            {
                // not in a domain
                ctx = new PrincipalContext(ContextType.Machine);
            }
            var up = UserPrincipal.FindByIdentity(ctx, username);
            if (up != null)
            {
                PrincipalSearchResult<Principal> authGroups = up.GetAuthorizationGroups();
                return authGroups.Any(principal =>
                                      principal.Sid.IsWellKnown(WellKnownSidType.BuiltinAdministratorsSid) ||
                                      principal.Sid.IsWellKnown(WellKnownSidType.AccountDomainAdminsSid) ||
                                      principal.Sid.IsWellKnown(WellKnownSidType.AccountAdministratorSid) ||
                                      principal.Sid.IsWellKnown(WellKnownSidType.AccountEnterpriseAdminsSid) ||
                                      principal.Sid.IsWellKnown(WellKnownSidType.BuiltinSystemOperatorsSid)
                                      );
            }
            return false;
        }
    }
}
