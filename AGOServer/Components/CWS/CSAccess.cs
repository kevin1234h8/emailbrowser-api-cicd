
using AGOServer.Components.Models.OpenText;
using AGOServer.Components.REST;
using CSAccess_Local;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using NLog;
using OpenText.Livelink.Service.Core;
using OpenText.Livelink.Service.DocMan;
using OpenText.Livelink.Service.MemberService;
using OpenText.Livelink.Service.SearchServices;
using RestSharp;
using Swashbuckle.Swagger;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI.WebControls;
using System.Xml.Linq;
using Version = OpenText.Livelink.Service.DocMan.Version;

namespace AGOServer
{
    public class CSAccess
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly LocalCSAccess CS = new LocalCSAccess();
        private static readonly RESTCSAccess RESTAccess = new RESTCSAccess();
        private static readonly CSRestAccess CSRestAccess = new CSRestAccess();
        private static readonly OTDSRestAccess OTDSRestAccess = new OTDSRestAccess();

        /// <summary>
        /// The following method validates the OTDS token in the 'otsession' cookie
        /// </summary>
        /// <param name="request">The incoming request containing the 'otsession' cookie</param>
        /// <param name="userNameToImpersonate">Username requested to impersonate</param>
        /// <returns>Whether the impersonation is valid</returns>
        public static async Task<bool> ValidateOTSession(HttpRequestMessage request, string userNameToImpersonate)
        {
            bool bypassSessionValidation = Properties.Settings.Default.BypassSessionValidation;

            if(bypassSessionValidation)
            {
                return true;
            }

            string otsessionToken = "";
            var cookies = request.Headers.GetCookies();

            foreach (var cookie in cookies)
            {
                var otsessionCookie = cookie.Cookies.FirstOrDefault(c => c.Name == "otsession");

                logger.Info("session");
                logger.Info(otsessionCookie);

                if (otsessionCookie != null)
                {
                    otsessionToken = otsessionCookie.Value.ToString();
                    break;
                }
            }

            //logger.Info($"Validating Token to impersonate ({userNameToImpersonate}): {otsessionToken}");
            bool isAuthenticated = false;

            if (string.IsNullOrEmpty(otsessionToken))
            {
                return false;
            }

            string currentUserApiUrl = Properties.Settings.Default.OTDS_REST_API + "/currentuser";

            if (!string.IsNullOrEmpty(otsessionToken))
            {
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    string loginName = "";

                    if (otsessionToken.StartsWith("b64."))
                    {
                        string base64Token = HttpUtility.UrlDecode(otsessionToken).Substring(4);
                        string decodedObjStr = ASCIIEncoding.UTF8.GetString(Convert.FromBase64String(base64Token));

                        Dictionary<string, object> userInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(decodedObjStr);

                        string OTCSTicket = userInfo["ott"].ToString();
                        V1CurrentUserResponse response = CSRestAccess.GetUserInfo(userNameToImpersonate, OTCSTicket);

                        if (response != null)
                        {
                            loginName = response.data["name"].ToString();
                        }
                    }
                    else
                    {
                        OTDSCurrentUserResponse response = OTDSRestAccess.GetCurrentUser(userNameToImpersonate, otsessionToken);

                        if (response != null)
                        {
                            loginName = response.user["name"].ToString();
                        }

                        #region oldCode
                        /*ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                        using (HttpClient client = new HttpClient())
                        {
                            // Set the bearer token in the Authorization header
                            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", otsessionToken);

                            // Send the GET request
                            HttpResponseMessage response = await client.GetAsync(currentUserApiUrl);

                            if (response.IsSuccessStatusCode)
                            {
                                string responseText = await response.Content.ReadAsStringAsync();
                                Dictionary<string, object> jsonResponse = JsonConvert.DeserializeObject<Dictionary<string, object>>(responseText);

                                if (jsonResponse.ContainsKey("user"))
                                {
                                    try
                                    {
                                        string loginName = (JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonResponse["user"].ToString()))["name"].ToString();
                                        isAuthenticated = (loginName.Equals(userNameToImpersonate, StringComparison.InvariantCultureIgnoreCase));
                                    }
                                    catch (Exception ex)
                                    {
                                        logger.Info($"Failed to authenticate the request Response received: {responseText}");
                                    }
                                }
                            }
                            else
                            {
                                logger.Info($"Failed to send out the API to verify the token");
                            }
                        }*/
                        #endregion
                    }

                    isAuthenticated = (loginName.Equals(userNameToImpersonate, StringComparison.InvariantCultureIgnoreCase));
                }
                catch (WebException ex)
                {
                    if (ex.Response is HttpWebResponse errorResponse)
                    {
                        HttpStatusCode errorCode = errorResponse.StatusCode;
                        logger.Error($"(Status Code: {errorCode.ToString()}) Not authorised to impersonate as {userNameToImpersonate}", ex);
                    }
                }
                catch (Exception exception)
                {
                    logger.Error($"Something went wrong when validating the token: {exception.Message}\r\n{exception.StackTrace}");
                }
            }

            return isAuthenticated;
        }
        public static (bool, string) CurrentUserGroupsCookie(HttpRequestMessage request, string userName) 
        {
            bool hasCookie = false;
            string encryptedUsergroup = "";

            var cookies = request.Headers.GetCookies();

            foreach (var cookie in cookies)
            {
                var otsessionCookie = cookie.Cookies.FirstOrDefault(c => c.Name == "ebid");

                if (otsessionCookie != null)
                {
                    hasCookie = true;
                    break;
                }
            }

            if (!hasCookie)
            {
                string usergroups = AGODataAccess.GetUserGroups(userName);
                encryptedUsergroup = SecureInfo.writeSensitiveInfo(usergroups);
                logger.Info($"GetUserGroups: {usergroups}({encryptedUsergroup})");
            }

            /*return (hasCookie, encryptedUsergroup);*/
            return (hasCookie, encryptedUsergroup);
        }

        private static string GetUserNameToImpersonate(string userNameToImpersonate)
        {
            return SessionAccess.GetUserNameForImpersonation(userNameToImpersonate);
        }
        public static User GetCurrentUser(string userNameToImpersonate)
        {
            return CS.GetCSUserByLoginName(GetUserNameToImpersonate(userNameToImpersonate));
        }
        public static Node[] ListNodes(string userNameToImpersonate, long nodeId, bool partialData)
        {
            return CS.ListNodes(GetUserNameToImpersonate(userNameToImpersonate), nodeId, partialData);
        }
        internal static bool IsNodeHidden(Node node)
        {
            return CS.IsNodeHidden(node);
        }
        internal static Node CreateFolder(string userNameToImpersonate, long parentID, string name, string comment, Metadata metadata)
        {
            return CS.CreateFolder(GetUserNameToImpersonate(userNameToImpersonate), parentID, name, comment, metadata);
        }
        public static Node GetNode(string userNameToImpersonate, long nodeId)
        {
            return CS.GetNode(GetUserNameToImpersonate(userNameToImpersonate), nodeId);
        }
        internal static FileAtts GetNodeLatestVersion(string userNameToImpersonate, long id, Stream outStream)
        {
            return CS.GetNodeLatestVersion(GetUserNameToImpersonate(userNameToImpersonate), id, outStream);
        }
        internal static SingleSearchResponse Search_Fuzzy(string userNameToImpersonate, string keywords, long folderNodeID, 
            long? folderNodeID2, long? folderNodeID3, long? folderNodeID4,
            int firstResultToRetrieve, int numResultsToRetrieve, string[] ResultColumns, string sortByColumn, string sortDirection,
            string OTEmailReceivedDate_From, string OTEmailReceivedDate_To, SearchFilterByAttachment searchFilterByAttachment,
            string OTEmailSubject, string OTEmailFrom, string OTEmailTo)
        {
            return CS.Search_Fuzzy(GetUserNameToImpersonate(userNameToImpersonate), keywords, folderNodeID, folderNodeID2, folderNodeID3, folderNodeID4,
                firstResultToRetrieve, numResultsToRetrieve, ResultColumns, sortByColumn, sortDirection,
             OTEmailReceivedDate_From, OTEmailReceivedDate_To, searchFilterByAttachment,
             OTEmailSubject,  OTEmailFrom,  OTEmailTo, OTSubType: "749");
        }
        public static Node CreateNodeAsAdmin(long parentID, string name, string type, string fileName, byte[] contents)
        {
            return CS.CreateNodeAsAdmin(parentID, name, type, fileName, contents);
        }
        public static NodeRights GetNodeRights(long nodeID)
        {
            return CS.GetNodeRights(nodeID);
        }
        public static void SetNodeRights(long nodeID, NodeRights nodeRights)
        {
            CS.SetNodeRights(nodeID, nodeRights);
        }
        public static Node GetNodeByNameAsAdmin(long parentNodeId, string name)
        {
            return CS.GetNodeByNameAsAdmin(parentNodeId, name);
        }
        public static Node CreateFolderAsAdmin(long parentID, string name, string comment, Metadata metadata)
        {
            return CS.CreateFolderAsAdmin(parentID, name, comment, metadata);
        }
        public static NodePageResult ListNodesByPage(string userNameToImpersonate, long parentID, int pageNumber, int pageSize, bool containersOnly, string[] includeTypes)
        {
            return CS.ListNodesByPage(GetUserNameToImpersonate(userNameToImpersonate), parentID, pageNumber, pageSize, containersOnly, includeTypes);
        }
        public static Rest_ListNodeResult ListNodesByPage_Rest(string userNameToImpersonate, long containerID, string sortedBy, string sortDirection, int pageSize, int page)
        {
            return CSRestAccess.ListEmailsByPage(GetUserNameToImpersonate(userNameToImpersonate), containerID, sortedBy, sortDirection, pageSize, page);
        }

        public static Rest_ListNodeResultVer2 ListNodesByPage_RestVer2(string userNameToImpersonate, long containerID, string sortedBy, string sortDirection, int pageSize, int page)
        {
            return CSRestAccess.ListEmailsByPageVer2(GetUserNameToImpersonate(userNameToImpersonate), containerID, sortedBy, sortDirection, pageSize, page);
        }
        public static Node GetFolderNodeByPath(string userNameToImpersonate, long nodeId, string[] path)
        {
            return CS.GetNodeByPath(GetUserNameToImpersonate(userNameToImpersonate), nodeId, path);
        }

        public static OpenTextV2SearchResponse Rest_Search_Fuzzy(Dictionary<string, string> searchOptions, string userNameToImpersonate)
        {
            return CSRestAccess.SearchV2(searchOptions, userNameToImpersonate);
        }
        public static OpenTextV2SearchResponse Rest_Search_Form(Dictionary<string, string> searchOptions, string userNameToImpersonate)
        {
            return CSRestAccess.SearchForm(searchOptions, userNameToImpersonate);
        }
    }
}