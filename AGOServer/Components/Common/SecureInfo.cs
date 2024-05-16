using CryptographyX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class SecureInfo
    {
        public static string getSensitiveInfo(string secureFileName)
        {
            string credentialsDirectory = Properties.Settings.Default.SecureCredentialsPath;
            string AESKeyFilePath = Path.Combine(credentialsDirectory, Properties.Settings.Default.SecureAESKey_Filename);
            string secureFilePath = Path.Combine(credentialsDirectory, secureFileName);
            return Cryptography.ReadSensitiveData(secureFilePath, AESKeyFilePath);
        }

        public static string readSensitiveInfo(string encrypted)
        {
            string credentialsDirectory = Properties.Settings.Default.SecureCredentialsPath;
            string AESKeyFilePath = Path.Combine(credentialsDirectory, Properties.Settings.Default.SecureAESKey_Filename);
            string passPhrase = File.ReadAllText(AESKeyFilePath);
            return Cryptography.Decrypt(encrypted, passPhrase);
        }

        public static string writeSensitiveInfo(string value)
        {
            string credentialsDirectory = Properties.Settings.Default.SecureCredentialsPath;
            string AESKeyFilePath = Path.Combine(credentialsDirectory, Properties.Settings.Default.SecureAESKey_Filename);
            string passPhrase = File.ReadAllText(AESKeyFilePath);
            return Cryptography.Encrypt(value, passPhrase);
        }
    }
}