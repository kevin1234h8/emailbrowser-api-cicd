using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using System.Web;

namespace AGOServer
{
    public class SessionAccess
    {
        public static SessionUser GetCurrentSessionUser()
        {
            SessionUser user = new SessionUser();
            
            user.LogonUserFullName = SessionAccess.GetLogonUserIdentityFullName();
            user.LogonUserSID = SessionAccess.GetLogonUserIdentitySID();
            user.IPAddress = SessionAccess.GetClientIPAddress();
            user.MAC = SessionAccess.GetClientMACAddress(user.IPAddress, HttpContext.Current.Request.IsLocal);
            return user;
        }

        public static string GetUserNameForImpersonation(string impersonateUserName)
        {
            bool UseWindowsAuthentication = Properties.Settings.Default.UseWindowsAuthentication;

            
            if (UseWindowsAuthentication)
            {
                impersonateUserName = GetLogonUserIdentityFullName();
            }

            bool excludeDomainFromName = Properties.Settings.Default.ExcludeDomainFromImpersonationUserName;
            if (excludeDomainFromName)
            {
                impersonateUserName = impersonateUserName.Substring(impersonateUserName.LastIndexOf('\\') + 1);
            }

            return impersonateUserName;
        }

        public static string GetLogonUserIdentityFullName()
        {
            string userIdentityName = "";
            try
            {
                userIdentityName = HttpContext.Current.Request.LogonUserIdentity.Name;
            }
            catch (Exception)
            {
            }
            return userIdentityName;
        }
        public static string GetLogonUserIdentitySID()
        {
            string sid = "";
            try
            {
                sid = HttpContext.Current.Request.LogonUserIdentity.User.AccountDomainSid.Value;
            }
            catch (Exception)
            {
            }
            return sid;
        }
        public static string GetClientIPAddress()
        {
            string clientIpAddress = "";
            try
            {
                clientIpAddress = System.Web.HttpContext.Current.Request.UserHostAddress;
            }
            catch (Exception)
            {
            }
            return clientIpAddress;
        }
        public static string GetClientMACAddress(string clientIpAddress, bool isLocal)
        {
            string MAC = "";
            try
            {
                MACRessolver ressolver = new MACRessolver();
                MAC = ressolver.GetMacByClientIp(clientIpAddress, isLocal);
                ressolver = null;
            }
            catch (Exception)
            {
            }
            if(string.IsNullOrEmpty(MAC)==false)
            {
                MAC = Regex.Replace(MAC, @"\s+", "");
            }
            return MAC;
        }

    }
}