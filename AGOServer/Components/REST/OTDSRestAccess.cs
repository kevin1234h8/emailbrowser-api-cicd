using AGOServer.Components.Models.OpenText;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;

namespace AGOServer.Components.REST
{
    public class OTDSRestAccess
    {
        private Logger logger = LogManager.GetCurrentClassLogger();

        #region Common Functions
        public RestClient GetRestClient()
        {
            return new RestClient(Properties.Settings.Default.OTDS_REST_API);
        }

        private NetworkCredential GetCSAdminUserNameAndPassword()
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

        private IRestResponse SendRestSharpRequest(RestRequest request, string userNameToImpersonate = null)
        {
            var client = GetRestClient();
            client.AddDefaultHeader("Bearer", GetAuthToken(userNameToImpersonate));
            IRestResponse response = client.Execute(request);

            return response;
        }

        private IRestResponse SendRestSharpRequestWithToken(RestRequest request, string token, string userNameToImpersonate = null)
        {
            var client = GetRestClient();
            //logger.Trace($"Sending Request to {client.BaseUrl}/{request.Resource} for {userNameToImpersonate} with Token: {token}");
            client.AddDefaultHeader("Authorization", $"Bearer {token}");
            IRestResponse response = client.Execute(request);

            return response;
        }

        private string GetAuthToken(string userNameToImpersonate = null)
        {
            string token = null;
            string userName = "";
            try
            {
                NetworkCredential credential = GetCSAdminUserNameAndPassword();
                userName = credential.UserName;
                token = RestSharp_Authenticate(userName, credential.Password, userNameToImpersonate);
                logger.Info("REST token retrieved for: " + userName + " is: " + token);
                credential = null;
            }
            catch (System.Exception ex)
            {
                logger.Error("Failed to authenticate as " + userName + " because " + ex.Message + ex.StackTrace, ex);
            }
            return token;
        }

        public string RestSharp_Authenticate(string username, string password, string userNameToImpersonate = null)
        {
            string token = "";
            try
            {
                var client = GetRestClient();
                //TODO: have to be v1
                var request = new RestRequest("authentication/credentials", Method.POST);
                request.AddParameter("username", username); // adds to POST or URL querystring based on Method
                request.AddParameter("password", password); // replaces matching token in request.Resource
                //request.AddParameter("otdsauth", "no-sso");

                if (username.Equals(userNameToImpersonate, StringComparison.OrdinalIgnoreCase) == false)
                {
                    request.AddParameter("impersonate_username", userNameToImpersonate);
                }

                // execute the request
                IRestResponse response = client.Execute(request);

                if (string.IsNullOrWhiteSpace(response.Content) == false)
                {
                    if (response.StatusCode == HttpStatusCode.OK)
                    {

                        JObject json = JObject.Parse(response.Content);
                        if (json != null && json.ContainsKey("ticket"))
                        {
                            token = json.SelectToken("ticket").ToString();
                        }
                        else
                        {
                            logger.Error("can't find ticket in response:" + response.Content);
                        }
                    }
                    else
                    {
                        logger.Error("failed to get Token , this is the response:" + response.Content);
                    }
                }
                else
                {
                    logger.Error("failed to get Token response content is null or empty ");
                }

            }
            catch (Exception ex)
            {
                logger.Error("failed to get Token " + ex.Message + ex.StackTrace, ex);
            }
            return token;
        }
        #endregion

        public OTDSCurrentUserResponse GetCurrentUser(string userNameToImpersonate, string token = null)
        {
            OTDSCurrentUserResponse result = null;
            
            try
            {
                string resource = "currentuser";
                var request = new RestRequest(resource, Method.GET);

                request.Timeout = 20 * 60 * 1000;
                IRestResponse response = (!string.IsNullOrEmpty(token)) ? SendRestSharpRequestWithToken(request, token, userNameToImpersonate) : SendRestSharpRequest(request, userNameToImpersonate);
                string content = response.Content;
                //logger.Trace($"Current User Result: {content}");

                result = JsonConvert.DeserializeObject<OTDSCurrentUserResponse>(content);
            }
            catch (Exception ex)
            {
                logger.Error($"Something went from trying to fetch data from the API: {ex.Message}\r\n{ex.StackTrace}");
            }

            return result;
        }
    }
}