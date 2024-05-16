using AGOServer.Components.Models.OpenText;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using NLog.Fluent;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using WebGrease.Activities;

namespace AGOServer.Components.REST
{
    public class CSRestAccess
    {
        private Logger logger = LogManager.GetCurrentClassLogger();

        public CSRestAccess() { }

        #region Common Functions
        public RestClient GetRestClient()
        {
            return new RestClient(Properties.Settings.Default.CS_REST_URL);
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
            client.AddDefaultHeader("OTCSTicket", GetAuthToken(userNameToImpersonate));
            IRestResponse response = client.Execute(request);

            return response;
        }

        private IRestResponse SendRestSharpRequestWithToken(RestRequest request, string token, string userNameToImpersonate = null)
        {
            var client = GetRestClient();
            client.AddDefaultHeader("OTCSTicket", token);
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
                var request = new RestRequest("v1/auth", Method.POST);
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

        private Rest_Node ParseNode(JToken obj)
        {

            Rest_Node node = new Rest_Node();
            if (ParseLongValue(obj, "id", out long id))
            {
                node.ID = id;
            }
            if (ParseStringValue(obj, "name", out string name))
            {
                node.Name = name;
            }
            if (ParseLongValue(obj, "parent_id", out long parent_id))
            {
                node.ParentID = parent_id;
            }
            if (ParseLongValue(obj, "size", out long size))
            {
                node.Size = size;
            }
            if (ParseIntValue(obj, "type", out int type))
            {
                node.Type = type;
            }
            if (ParseBoolValue(obj, "perm_see_contents", out bool perm_see_contents))
            {
                logger.Info($"perm_see_contents {perm_see_contents}");
                node.PermSeeContents = perm_see_contents;
            }
            logger.Info($"node from ParseNode {node}");
            logger.Info($"node.PermSeeContents from ParseNode {node.PermSeeContents}");
            return node;
        }

        private Rest_NodeVer2 ParseNodeVer2(JToken prop, JToken emailProp)
        {
            Rest_NodeVer2 node = new Rest_NodeVer2();
            if (ParseLongValue(prop, "id", out long id))
            {
                node.Id = id;
            }
            if (ParseStringValue(prop, "name", out string name))
            {
                node.Name = name;
            }
            if (ParseLongValue(prop, "parent_id", out long parentId))
            {
                node.ParentId = parentId;
            }
            if (ParseLongValue(prop, "size", out long size))
            {
                node.Size = size;
            }
            if (ParseIntValue(prop, "type", out int type))
            {
                node.Type = type;
            }
            if (ParseStringValue(emailProp, "subject", out string emailSubject))
            {
                node.EmailSubject = emailSubject;
            }
            if (ParseStringValue(emailProp, "to", out string emailTo))
            {
                node.EmailTo = emailTo;
            }
            if (ParseStringValue(emailProp, "cc", out string emailCc))
            {
                node.EmailCc = emailCc;
            }
            if (ParseStringValue(emailProp, "from", out string emailFrom))
            {
                node.EmailFrom = emailFrom;
            }
            if (ParseDateTimeValue(emailProp, "sentdate", out DateTime sentDate))
            {
                node.SentDate = sentDate;
            }
            if (ParseDateTimeValue(emailProp, "receiveddate", out DateTime receivedDate))
            {
                node.ReceivedDate = receivedDate;
            }
            if (ParseIntValue(emailProp, "hasattachments", out int hasAttachments))
            {
                node.HasAttachments = hasAttachments;
            }
            if (ParseStringValue(emailProp, "conversationid", out string conversationId))
            {
                node.ConversationId = conversationId;
            }

            return node;
        }

        private bool ParseStringValue(JToken obj, string key, out string result)
        {
            bool success = false;
            result = string.Empty;
            try
            {
                var child = obj[key];
                if (child != null)
                {
                    result = child.Value<string>();
                    success = true;
                }
            }
            catch (Exception ex)
            {
                logger.Warn("failed to parse key[" + key + "]json obj[" + obj.ToString() + "] reason:" + ex.Message + ex.StackTrace);
            }
            return success;
        }

        private bool ParseBoolValue(JToken obj, string key, out bool result)
        {
            bool success = false;
            result = false;
            try
            {
                var child = obj[key];
                if (child != null)
                {
                    result = child.Value<bool>();
                    success = true;
                }
            }
            catch (Exception ex)
            {
                logger.Warn("failed to parse key[" + key + "]json obj[" + obj.ToString() + "] reason:" + ex.Message + ex.StackTrace);
            }
            return success;
        }
        private bool ParseLongValue(JToken obj, string key, out long result)
        {
            bool success = false;
            result = 0;
            try
            {
                var child = obj[key];
                if (child != null)
                {
                    result = child.Value<long>();
                    success = true;
                }
            }
            catch (Exception ex)
            {
                logger.Warn("failed to parse key[" + key + "]json obj[" + obj.ToString() + "] reason:" + ex.Message + ex.StackTrace);
            }
            return success;
        }
        private bool ParseDateTimeValue(JToken obj, string key, out DateTime result)
        {
            bool success = false;
            result = DateTime.MinValue;
            try
            {
                var child = obj[key];
                if (child != null)
                {
                    result = child.Value<DateTime>();
                    success = true;
                }
            }
            catch (Exception ex)
            {
                logger.Warn("failed to parse key[" + key + "]json obj[" + obj.ToString() + "] reason:" + ex.Message + ex.StackTrace);
            }
            return success;
        }
        private bool ParseIntValue(JToken obj, string key, out int result)
        {
            bool success = false;
            result = 0;
            try
            {
                var child = obj[key];
                if (child != null)
                {
                    result = child.Value<int>();
                    success = true;
                }
            }
            catch (Exception ex)
            {
                logger.Warn("failed to parse key[" + key + "]json obj[" + obj.ToString() + "] reason:" + ex.Message + ex.StackTrace);
            }
            return success;
        }
        #endregion

        public Rest_ListNodeResult ListEmailsByPage(string userNameToImpersonate, long id, string sortedBy, string sortDirection, int limit, int page = 1)
        {
            Rest_ListNodeResult result = new Rest_ListNodeResult();
            List<Rest_Node> nodes = new List<Rest_Node>();
            string resource = "v1/nodes/" + id + "/nodes";
            var request = new RestRequest(resource, Method.GET);

            try
            {
                // add files to upload (works with compatible verbs)
                request.Timeout = 20 * 60 * 1000;
                //request.AlwaysMultipartFormData = true;
                //request.AddParameter("where_type", -3);//get non containers. cannot be used with facets
                request.AddParameter("where_type", 749);//get non containers. cannot be used with facets
                //request.AddParameter("where_facet", "4096:749");//4096 is content type, 749 is email. this lets us filter by email
                request.AddParameter("extra", "False");//disable getting extra info
                request.AddParameter("page", page);
                request.AddParameter("limit", limit);
                if (string.IsNullOrWhiteSpace(sortedBy) == false)
                {
                    request.AddParameter("sort", sortDirection + "_" + "wnf_" + sortedBy);
                }

                // execute the request
                IRestResponse response = SendRestSharpRequest(request, userNameToImpersonate);
                string content = response.Content;// raw content as string

                if (string.IsNullOrWhiteSpace(response.Content) == false)
                {
                    JObject json = null;
                    try
                    {
                        json = JObject.Parse(response.Content);
                    }
                    catch (Exception ex)
                    {
                        logger.Error("response JSON is invalid [" + response.Content + "] reason: " + ex.Message + ex.StackTrace);
                    }

                    if (json != null)
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            if (json.ContainsKey("data"))
                            {
                                var data = json.SelectToken("data");
                                if (data != null)
                                {
                                    foreach (var obj in data)
                                    {
                                        var node = ParseNode(obj);
                                        nodes.Add(node);
                                    }
                                }
                            }
                            if (ParseStringValue(json, "sort", out string actualSortedBy))
                            {
                                result.ActualSortedBy = actualSortedBy;
                            }
                            if (ParseIntValue(json, "page_total", out int page_total))
                            {
                                result.TotalPage = page_total;
                            }
                            if (ParseIntValue(json, "total_count", out int total_count))
                            {
                                result.TotalCount = total_count;
                            }
                            if (ParseIntValue(json, "range_min", out int range_min))
                            {
                                result.RangeMin = range_min;
                            }
                            if (ParseIntValue(json, "range_max", out int range_max))
                            {
                                result.RangeMax = range_max;
                            }
                        }
                        else
                        {
                            if (json.ContainsKey("error"))
                            {
                                string error = json.SelectToken("error").ToString();
                                if (string.IsNullOrWhiteSpace(error) == false)
                                {
                                    throw new Exception(error);
                                }
                            }
                            logger.Warn("unexpected response content : " + response.Content);
                        }
                    }
                    else
                    {
                        logger.Warn("response JSON is null");

                    }
                }
                else
                {
                    logger.Warn("response is null");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            result.Nodes = nodes;
            return result;
        }

        public Rest_ListNodeResultVer2 ListEmailsByPageVer2(string userNameToImpersonate, long id, string sortedBy, string sortDirection, int limit, int page)
        {
            Rest_ListNodeResultVer2 result = new Rest_ListNodeResultVer2();
            List<Rest_NodeVer2> nodes = new List<Rest_NodeVer2>();
            string resource = "v2/nodes/" + id + "/nodes";
            var request = new RestRequest(resource, Method.GET);
            try
            {
                request.Timeout = 20 * 60 * 1000;
                request.AddParameter("where_type", 749); // Filter on node type. 
                request.AddParameter("page", page); // Page number.
                request.AddParameter("limit", limit); // Page size.
                if (string.IsNullOrWhiteSpace(sortedBy) == false)
                {
                    request.AddParameter("sort", sortDirection + "_wnf_" + sortedBy);
                }

                // Execute the request.                
                IRestResponse response = SendRestSharpRequest(request, userNameToImpersonate);
                string content = response.Content;

                if (string.IsNullOrWhiteSpace(response.Content) == false)
                {
                    JObject json = null;
                    try
                    {
                        json = JObject.Parse(response.Content);
                    }
                    catch (Exception ex)
                    {
                        logger.Error("response JSON is invalid [" + response.Content + "] reason: " + ex.Message + ex.StackTrace);
                    }

                    if (json != null)
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            if (json.ContainsKey("results"))
                            {
                                var objectResults = json.SelectToken("results");
                                if (objectResults.Count() > 0)
                                {
                                    foreach (var obj in objectResults)
                                    {
                                        var data = obj.SelectToken("data");
                                        var prop = data.SelectToken("properties");
                                        var emailProp = data.SelectToken("emailproperties");

                                        var node = ParseNodeVer2(prop, emailProp);
                                        nodes.Add(node);
                                    }
                                }
                            }

                            if (json.ContainsKey("collection"))
                            {
                                var collection = json.SelectToken("collection");
                                var paging = collection.SelectToken("paging");
                                var sorting = collection.SelectToken("sorting");
                                var sort = sorting.SelectToken("sort").FirstOrDefault(x => x.SelectToken("key").ToString() == "sort");

                                if (ParseStringValue(sort, "value", out string actualSortedBy))
                                {
                                    result.ActualSortedBy = actualSortedBy;
                                }
                                if (ParseIntValue(paging, "page_total", out int pageTotal))
                                {
                                    result.TotalPage = pageTotal;
                                }
                                if (ParseIntValue(paging, "total_count", out int totalCount))
                                {
                                    result.TotalCount = totalCount;
                                }
                                if (ParseIntValue(paging, "range_min", out int rangeMin))
                                {
                                    result.RangeMin = rangeMin;
                                }
                                if (ParseIntValue(paging, "range_max", out int rangeMax))
                                {
                                    result.RangeMax = rangeMax;
                                }
                                if (ParseIntValue(paging, "page", out int pageNumber))
                                {
                                    result.PageNumber = pageNumber;
                                }
                                if (ParseIntValue(paging, "limit", out int pageSize))
                                {
                                    result.PageSize = pageSize;
                                }
                            }
                        }
                        else
                        {
                            if (json.ContainsKey("error"))
                            {
                                string error = json.SelectToken("error").ToString();
                                if (string.IsNullOrWhiteSpace(error) == false)
                                {
                                    throw new Exception(error);
                                }
                            }
                            logger.Warn("Unexpected response content: " + response.Content);
                        }
                    }
                    else
                    {
                        logger.Warn("Response JSON is null.");
                    }
                }
                else
                {
                    logger.Warn("Response is null.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            result.Nodes = nodes;
            return result;
        }

        public OpenTextV2SearchResponse SearchV2(Dictionary<string, string> searchOptions, string userNameToImpersonate = null)
        {
            OpenTextV2SearchResponse searchResponse = new OpenTextV2SearchResponse();
            string resource = "v2/search";
            var request = new RestRequest(resource, Method.POST);

            try
            {
                List<string> queryList = new List<string>();
                StringBuilder queryStrBuilder = new StringBuilder();
                int page = 1, limit = 30;
                int folderID = -1, folderID2 = -1, folderID3 = -1, folderID4 = -1;
                string sortColumn = searchOptions["SortColumn"], sortDirection = searchOptions["SortDirection"];
                string keywords = searchOptions["keywords"], emailSubjectKeywords = searchOptions["OTEmailSubject"];
                string emailFrom = searchOptions["OTEmailFrom"], emailTo = searchOptions["OTEmailTo"];
                string emailSentFromDate = searchOptions["OTEmailSentDate_Start"], emailSentToDate = searchOptions["OTEmailSentDate_End"], emailReceiveFromDate = searchOptions["OTEmailReceivedDate_Start"], emailReceiveToDate = searchOptions["OTEmailReceivedDate_End"];
                string attachmentFileName = searchOptions["AttachmentName"];
                string cache_id = searchOptions["CacheID"];
                string filterByAttachment = searchOptions["SearchByAttachment"];

                int.TryParse(searchOptions["page"].ToString(), out page);
                int.TryParse(searchOptions["limit"].ToString(), out limit);
                int.TryParse(searchOptions["FolderNodeID"].ToString(), out folderID);
                int.TryParse(searchOptions["FolderNodeID2"].ToString(), out folderID2);
                int.TryParse(searchOptions["FolderNodeID3"].ToString(), out folderID3);
                int.TryParse(searchOptions["FolderNodeID4"].ToString(), out folderID4);

                if (!string.IsNullOrEmpty(keywords) && string.IsNullOrEmpty(attachmentFileName))
                {
                    if (keywords.Count(c => c == '"') % 2 == 0)
                        queryList.Add(keywords);
                    else
                        queryList.Add(string.Format(Properties.Settings.Default.KeywordSearchPattern, keywords));
                }

                //if (!string.IsNullOrEmpty(keywords) && !string.IsNullOrEmpty(emailSubjectKeywords))
                //{
                //    queryList.Add($"({keywords} OR OTEmailSubject: {emailSubjectKeywords})");
                //}
                //else if (!string.IsNullOrEmpty(keywords) && string.IsNullOrEmpty(emailSubjectKeywords))
                //{
                //    queryList.Add($"({keywords} OR OTEmailSubject: {keywords})");
                //}
                //else if (string.IsNullOrEmpty(keywords) && !string.IsNullOrEmpty(emailSubjectKeywords))
                //{
                //    queryList.Add($"({emailSubjectKeywords} OR OTEmailSubject: {emailSubjectKeywords})");
                //}

                if (!string.IsNullOrEmpty(emailSubjectKeywords))
                {
                    queryList.Add($"(OTEmailSubject: {emailSubjectKeywords})");
                }

                if (!string.IsNullOrEmpty(emailFrom))
                {
                    queryList.Add($"(OTEmailFrom: {emailFrom}* OR OTEmailSenderAddress: {emailFrom}* OR OTEmailSenderFullName: {emailFrom}*)");
                }

                if (!string.IsNullOrEmpty(emailTo))
                {
                    var recipients = emailTo.Split(';');
                    var OTEmailTo = string.Join(" & ", recipients.Select(x => $"OTEmailTo: {x}*"));
                    var OTEmailToAddress = string.Join(" & ", recipients.Select(x => $"OTEmailToAddress: {x}*"));
                    var OTEmailToFullName = string.Join(" & ", recipients.Select(x => $"OTEmailToFullName: {x}*"));
                    var OTEmailRecipientAddress = string.Join(" & ", recipients.Select(x => $"OTEmailRecipientAddress: {x}*"));
                    var OTEmailRecipientFullName = string.Join(" & ", recipients.Select(x => $"OTEmailRecipientFullName: {x}*"));
                    queryList.Add($"(({OTEmailTo}) OR ({OTEmailToAddress}) OR ({OTEmailToFullName}) OR ({OTEmailRecipientAddress}) OR ({OTEmailRecipientFullName}))");
                }

                if (!string.IsNullOrEmpty(emailSentFromDate) && !string.IsNullOrEmpty(emailSentToDate))
                {
                    //$"([qlregion \"OTEmailSentDate\"] >= {emailSentFromDate} AND [qlregion \"OTEmailSentDate\"] <= {emailSentFromDate})"
                    //$"(OTEmailSentDate (> {emailSentFromDate}) AND OTEmailSentDate (< {emailSentToDate}))"
                    queryList.Add($"([qlregion \"OTEmailSentDate\"] >= {emailSentFromDate} AND [qlregion \"OTEmailSentDate\"] <= {emailSentToDate})");
                }

                if (!string.IsNullOrEmpty(emailReceiveFromDate) && !string.IsNullOrEmpty(emailReceiveToDate))
                {
                    queryList.Add($"(OTEmailReceivedDate >= {emailReceiveFromDate} AND OTEmailReceivedDate <= {emailReceiveToDate})");
                }

                if (!string.IsNullOrEmpty(attachmentFileName))
                {
                    queryList.Add($"(({attachmentFileName}) AND [qlregion \"OTEmailHasAttachments\"] = \"true\")");
                }
                else if (!string.IsNullOrEmpty(filterByAttachment) && filterByAttachment == "1")
                {
                    queryList.Add($"([qlregion \"OTEmailHasAttachments\"] = \"true\")");
                }
                else if (!string.IsNullOrEmpty(filterByAttachment) && filterByAttachment == "0")
                {
                    queryList.Add($"([qlregion \"OTEmailHasAttachments\"] = \"false\")");
                }

                queryStrBuilder.Append(string.Join(" AND ", queryList));
                string query = $"{queryStrBuilder.ToString()}";
                logger.Trace($"Search Query: \"{query}\"\tOptions: {JsonConvert.SerializeObject(searchOptions)}");

                request.AddParameter("page", page);
                request.AddParameter("limit", limit);
                request.AddParameter("select", "{'OTEmailCC','OTEmailCCAddress','OTEmailCCFullName','OTEmailFrom','OTEmailHasAttachments','OTEmailReceivedDate','OTEmailRecipientAddress','OTEmailRecipientFullName','OTEmailSenderAddress','OTEmailSenderFullName','OTEmailSentDate','OTEmailSubject','OTEmailTo','OTEmailToAddress','OTEmailToFullName'}");
                request.AddParameter("lookfor", "complexquery");
                request.AddParameter("filter", "OTSubType: {749}");
                request.AddParameter("modifier", "relatedto");
                request.AddParameter("fields", "properties");
                request.AddParameter("where", queryStrBuilder.ToString());
                request.AddParameter("sort", $"{sortDirection}_{sortColumn}");
                //request.AddParameter("options", "{'highlight_summaries'}");

                if (folderID > 0)
                    request.AddParameter("location_id1", folderID);

                if (folderID2 > 0)
                    request.AddParameter("location_id2", folderID2);

                if (folderID3 > 0)
                    request.AddParameter("location_id3", folderID3);

                if (folderID4 > 0)
                    request.AddParameter("location_id4", folderID4);

                if (!string.IsNullOrEmpty(cache_id))
                {
                    request.AddParameter("cache_id", cache_id);
                }

                //logger.Info($"Search Request Parameters: {string.Join("|", request.Parameters.Select(p => $"{p.Name}: {p.Value.ToString()}"))}");
                request.Timeout = 20 * 60 * 1000;
                IRestResponse response = SendRestSharpRequest(request, userNameToImpersonate);
                string content = response.Content;
                //logger.Trace($"Search Result: {content}");

                searchResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<OpenTextV2SearchResponse>(content);
            }
            catch (Exception ex)
            {
                logger.Error($"Something went from trying to fetch data from the API: {ex.Message}\r\n{ex.StackTrace}");
            }

            return searchResponse;
        }

        public OpenTextV2SearchResponse SearchForm(Dictionary<string, string> searchOptions, string userNameToImpersonate = null)
        {
            OpenTextV2SearchResponse searchResponse = new OpenTextV2SearchResponse();
            string resource = "v2/search";
            var request = new RestRequest(resource, Method.POST);

            try
            {
                List<string> queryList = new List<string>();
                StringBuilder queryStrBuilder = new StringBuilder();
                int page = 1, limit = 30;
                int folderID = -1, folderID2 = -1, folderID3 = -1, folderID4 = -1;
                string sortColumn = searchOptions["SortColumn"], sortDirection = searchOptions["SortDirection"];
                string keywords = searchOptions["keywords"], emailSubjectKeywords = searchOptions["OTEmailSubject"];
                string emailFrom = searchOptions["OTEmailFrom"], emailTo = searchOptions["OTEmailTo"];
                string emailSentFromDate = searchOptions["OTEmailSentDate_Start"], emailSentToDate = searchOptions["OTEmailSentDate_End"], emailReceiveFromDate = searchOptions["OTEmailReceivedDate_Start"], emailReceiveToDate = searchOptions["OTEmailReceivedDate_End"];
                string attachmentFileName = searchOptions["AttachmentName"];

                int.TryParse(searchOptions["page"].ToString(), out page);
                int.TryParse(searchOptions["limit"].ToString(), out limit);
                int.TryParse(searchOptions["FolderNodeID"].ToString(), out folderID);
                int.TryParse(searchOptions["FolderNodeID2"].ToString(), out folderID2);
                int.TryParse(searchOptions["FolderNodeID3"].ToString(), out folderID3);
                int.TryParse(searchOptions["FolderNodeID4"].ToString(), out folderID4);

                request.AddParameter("FullText_value1", keywords);
                request.AddParameter("SystemAttributes_value1_ID", "");
                request.AddParameter("BrowseLivelink_value1_ID", searchOptions["FolderNodeID"].ToString());
                request.AddParameter("BrowseLivelink_value2_ID", searchOptions["FolderNodeID2"].ToString());
                request.AddParameter("EmailAttributes_Subject", emailSubjectKeywords);
                request.AddParameter("EmailAttributes_Participant", "");
                request.AddParameter("EmailAttributes_RecipientOpt", "true");
                request.AddParameter("EmailAttributes_SenderOpt", "true");
                request.AddParameter("EmailAttributes_ReceivedType", "anydate");

                if (!string.IsNullOrEmpty(emailSentFromDate) && !string.IsNullOrEmpty(emailSentToDate))
                {
                    string fromDate = $"{emailSentFromDate.Substring(0, 4)}-{emailSentFromDate.Substring(4, 2)}-{emailSentFromDate.Substring(6, 2)}";
                    string toDate = $"{emailSentToDate.Substring(0, 4)}-{emailSentToDate.Substring(4, 2)}-{emailSentToDate.Substring(6, 2)}";
                    logger.Trace($"From: {emailSentFromDate} ({((fromDate != null) ? fromDate.ToString() : "null")}) To: {emailSentToDate} ({((toDate != null) ? toDate.ToString() : "null")})");
                    request.AddParameter("EmailAttributes_SentType", "range");
                    request.AddParameter("EmailAttributes_SentType_DFrom", fromDate);
                    request.AddParameter("EmailAttributes_SentType_DTo", toDate);
                }
                else
                {
                    request.AddParameter("EmailAttributes_SentType", "anydate");
                }

                request.AddParameter("page", page);
                request.AddParameter("limit", limit);
                request.AddParameter("select", "{'OTEmailCC','OTEmailCCAddress','OTEmailCCFullName','OTEmailFrom','OTEmailHasAttachments','OTEmailReceivedDate','OTEmailRecipientAddress','OTEmailRecipientFullName','OTEmailSenderAddress','OTEmailSenderFullName','OTEmailSentDate','OTEmailSubject','OTEmailTo','OTEmailToAddress','OTEmailToFullName'}");
                request.AddParameter("lookfor", "complexquery");
                request.AddParameter("query_id", Properties.Settings.Default.SearchFormQueryID.ToString());
                request.AddParameter("modifier", "relatedto");
                request.AddParameter("fields", "properties");
                request.AddParameter("sort", $"{sortDirection}_{sortColumn}");
                request.AddParameter("options", "{'highlight_summaries'}");

                if (folderID > 0)
                    request.AddParameter("location_id1", folderID);

                if (folderID2 > 0)
                    request.AddParameter("location_id2", folderID2);

                if (folderID3 > 0)
                    request.AddParameter("location_id3", folderID3);

                if (folderID4 > 0)
                    request.AddParameter("location_id4", folderID4);

                request.Timeout = 20 * 60 * 1000;
                IRestResponse response = SendRestSharpRequest(request, userNameToImpersonate);
                string content = response.Content;
                //logger.Trace($"Search Result: {content}");

                searchResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<OpenTextV2SearchResponse>(content);
            }
            catch (Exception ex)
            {
                logger.Error($"Something went from trying to fetch data from the API: {ex.Message}\r\n{ex.StackTrace}");
            }

            return searchResponse;
        }

        public V1CurrentUserResponse GetUserInfo(string userNameToImpersonate, string token = null)
        {
            V1CurrentUserResponse result = null;

            try
            {
                string resource = "v1/auth";
                var request = new RestRequest(resource, Method.GET);

                request.Timeout = 20 * 60 * 1000;
                IRestResponse response = (!string.IsNullOrEmpty(token)) ? SendRestSharpRequestWithToken(request, token, userNameToImpersonate) : SendRestSharpRequest(request, userNameToImpersonate);
                string content = response.Content;
                //logger.Trace($"Current User Result: {content}");

                result = JsonConvert.DeserializeObject<V1CurrentUserResponse>(content);
            }
            catch (Exception ex)
            {
                logger.Error($"Something went from trying to fetch data from the API: {ex.Message}\r\n{ex.StackTrace}");
            }

            return result;
        }
    }
}