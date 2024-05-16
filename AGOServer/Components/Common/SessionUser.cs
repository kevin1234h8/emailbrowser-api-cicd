using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class SessionUser
    {
        string _LogonUserSID = "";
        string _IPAddress = "";
        string _MAC = "";
        string _LogonUserFullName = "";

        public string LogonUserSID { get => _LogonUserSID; set => _LogonUserSID = value; }
        public string IPAddress { get => _IPAddress; set => _IPAddress = value; }
        public string MAC { get => _MAC; set => _MAC = value; }
        public string LogonUserFullName { get => _LogonUserFullName; set => _LogonUserFullName = value; }
    }
}