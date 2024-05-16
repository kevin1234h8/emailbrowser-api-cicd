using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using OpenText.Livelink.Service.SearchServices;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Text;
using System.IO;
using OpenText.Livelink.Service.DocMan;
using OpenText.Livelink.Service.Core;
using Microsoft.Ajax.Utilities;
using CSAccess_Local;

namespace AGOServer
{
    public class LocalCSAccess : CSAccess_Local_API
    {
        #region overheads
        private Logger logger = LogManager.GetCurrentClassLogger();

        public override NetworkCredential GetCSAdminUserNameAndPassword()
        {
            NetworkCredential credential = new NetworkCredential();
            if (Properties.Settings.Default.UseSecureCredentials)
            {
                credential.UserName = SecureInfo.getSensitiveInfo(Properties.Settings.Default.SecureOTCSUsername_Filename);
                credential.Password = SecureInfo.getSensitiveInfo(Properties.Settings.Default.SecureOTCSSecret_Filename);
            }
            else
            {

                credential.UserName = Properties.Settings.Default.OTCS_Username;
                credential.Password = Properties.Settings.Default.OTCS_Password;
            }
            return credential;
        }

        public override void LogTrace(string text)
        {
            logger.Trace(text);
        }
        public override void LogInfo(string text)
        {
            logger.Info(text);
        }

        public override void LogDebug(string text)
        {
            logger.Debug(text);
        }
        public override void LogWarn(string text, Exception ex = null)
        {
            logger.Warn(text);
        }

        public override void LogError(string text, Exception ex)
        {
            logger.Error(text);
        }
        public override void LogFatal(string text, Exception ex = null)
        {
            logger.Fatal(text);
        }
        #endregion

        internal string Debug_GenerateSearchInfo(string userNameToImpersonate)
        {
            SearchService searchService = GetSearchService(userNameToImpersonate);
            string queryLanguageDesc = searchService.GetQueryLanguageDescription("eng", "Livelink Search API V1.1");//"V1 and V1.1
            string[] dataCollections = searchService.GetDataCollections();
            File.WriteAllText(@"C:\_work\emailbrowserapi\queryLanguageDesc2.txt", queryLanguageDesc);
            if (dataCollections != null)
            {
                for (int i = 0; i < dataCollections.Length; i++)
                {
                    File.WriteAllText(@"C:\_work\emailbrowserapi\dataCollections_" + i + ".txt", dataCollections[i]);
                    FieldInfo[] fieldInfos = searchService.GetFieldInfo("'" + dataCollections[i] + "'", null);
                    if (fieldInfos != null)
                    {
                        StringBuilder sb = new StringBuilder();

                        for (int j = 0; j < fieldInfos.Length; j++)
                        {
                            string fieldInfoStr = Prettify(fieldInfos[j]);
                            sb.AppendLine(fieldInfoStr);
                            sb.AppendLine(System.Environment.NewLine);
                        }
                        File.WriteAllText(@"C:\_work\emailbrowserapi\dataCollections_" + i + "_fieldInfo.txt", sb.ToString());

                    }
                }
            }
            return queryLanguageDesc;
        }
        private string Prettify(Object obj)
        {
            var jsonString = JsonConvert.SerializeObject(
           obj, Formatting.Indented,
           new JsonConverter[] { new StringEnumConverter() });
            return jsonString;
        }
    }
}