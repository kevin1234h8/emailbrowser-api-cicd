using System;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;

namespace AGOServer
{
    public class Utilx
    {
        public static string GetRelativeString(string longer,string shorter)
        {
            string result = "";
            if(shorter.Length <= longer.Length)
            {
                result = longer.Substring(shorter.Length);
            }
            return result;
        }
        public static void postFixWithForwardSlash(ref String theUri)
        {
            if (theUri[theUri.Length - 1] != '/')
            {
                theUri += '/';
            }
        }
        public static void removeLastForwardSlash(ref String theUri)
        {
            if (theUri[theUri.Length - 1] == '/')
            {
                theUri = theUri.Substring(0, theUri.Length - 1);
            }
        }
        private static string ToQueryString(NameValueCollection nvc)
        {
            var array = (from key in nvc.AllKeys
                         from value in nvc.GetValues(key)
                         select string.Format("{0}={1}", WebUtility.UrlEncode(key), WebUtility.UrlEncode(value)))
                .ToArray();
            return "?" + string.Join("&", array);
        }
        public static string GetSafeFilename(string filename,char replacementChar)
        {
            return string.Join(replacementChar.ToString(), filename.Split(Path.GetInvalidFileNameChars()));
        }
        public static bool ContainsIllegalCharacters(string filename)
        {
            string[] invalidChars = filename.Split(Path.GetInvalidFileNameChars());
            if(invalidChars.Length > 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
