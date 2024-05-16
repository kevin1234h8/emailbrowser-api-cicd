using RestSharp;
using SXCommonRESTCSAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using NLog;
using Newtonsoft.Json.Linq;
using NLog.Targets;
using System.Collections.Concurrent;

namespace AGOServer
{
    public class RESTCSAccess : CommonAPIRESTCSAccess
    {
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

        public override void LogError(string text, Exception ex)
        {
            logger.Error(text);
        }

        public override void LogInfo(string text)
        {
            logger.Info(text);
        }

        public override void LogWarn(string text, Exception ex = null)
        {
            logger.Warn(text);
        }

        public override RestClient GetRestClient()
        {
            return new RestClient(Properties.Settings.Default.CS_REST_URL);
        }        

        public Rest_ListNodeResult ListEmailsByPage(string userNameToImpersonate, long id, string sortedBy, string sortDirection, int limit, int page = 1)
        {
            Rest_ListNodeResult result = new Rest_ListNodeResult();
            List<Rest_Node> nodes = new List<Rest_Node>();
            string resource = "nodes/" + id + "/nodes";
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
                if(string.IsNullOrWhiteSpace(sortedBy)==false)
                {
                    request.AddParameter("sort", sortDirection + "_" + "wnf_" + sortedBy);
                }

                // execute the request
                IRestResponse response = MakeRestSharpRequest(request, userNameToImpersonate);
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
                        LogWarn("response JSON is invalid ["+response.Content +"] reason: "+ex.Message+ex.StackTrace);
                    }

                    if (json != null)
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            if (json.ContainsKey("data"))
                            {
                                var data = json.SelectToken("data");
                                if(data!=null)
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
                            LogWarn("unexpected response content : " + response.Content);
                        }
                    }
                    else
                    {
                        LogWarn("response JSON is null");

                    }
                }
                else
                {
                    LogWarn("response is null");
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
            string resource = "/v2/nodes/" + id + "/nodes";
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
                IRestResponse response = MakeRestSharpRequest(request, userNameToImpersonate);
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
                        LogWarn("response JSON is invalid [" + response.Content + "] reason: " + ex.Message + ex.StackTrace);
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
                            LogWarn("Unexpected response content: " + response.Content);
                        }
                    }
                    else
                    {
                        LogWarn("Response JSON is null.");
                    }
                }
                else
                {
                    LogWarn("Response is null.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            result.Nodes = nodes;
            return result;
        }

        private Rest_Node ParseNode(JToken obj)
        {
            Rest_Node node = new Rest_Node();
            if(ParseLongValue(obj, "id", out long id))
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
                logger.Warn("failed to parse key["+key+"]json obj["+obj.ToString()+"] reason:"+ex.Message+ex.StackTrace);
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
    }
}