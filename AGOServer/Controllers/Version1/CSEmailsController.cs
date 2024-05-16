using AGOServer.Components;
using CSAccess_Local;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using System.Web.Http.Results;

namespace AGOServer
{
    /// <summary>
    /// Content Server Emails
    /// </summary>
    /// <remarks>
    /// Emails that reside in CS
    /// </remarks>
    [RoutePrefix("api/v1/emails")]
    public class CSEmailsController : ApiController
    {

        /// <summary>
        /// Get the list of immediate child emails from a parent CS Folder
        /// </summary>
        /// <remarks>
        /// Return a List of Content Server Email informations
        /// </remarks>
        /// <param name="id">The Content Server ID of the parent folder</param>
        /// <param name="includeConversationID">Indicate if you want to retrieve the conversation ID for the emails</param>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <param name="sortedBy">
        /// Default is OTEmailReceivedDate.
        /// Available options:
        /// OTEmailFrom
        /// OTEmailSentDate
        /// OTEmailTo
        /// OTEmailReceivedDate
        /// OTEmailCC
        /// OTEmailHasAttachments
        /// OTEmailSubject
        /// </param>
        /// <param name="sortDirection">Direction of the sort. Either "asc" or "desc", default is "desc"</param>
        /// <returns></returns>
        /// 

        [System.Web.Http.HttpGet, System.Web.Http.Route("GetEmailBrowserSummarizerApiBaseUri")]
        public async Task<IHttpActionResult> GetEmailBrowserSummarizerApiBaseUri()
        {
            string emailBrowserSummarizerApiBaseUri = Properties.Settings.Default.EmailBrowserSummarizerApiBaseUri;
            var result = Request.CreateResponse(HttpStatusCode.OK, emailBrowserSummarizerApiBaseUri);

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }


        [ResponseType(typeof(ListEmailsResult))]
        [HttpGet, Route("list")]
        public async Task<IHttpActionResult> GetEmails(long id, bool includeConversationID = false, string userName = "", string sortedBy = "", string sortDirection = "desc")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                ListEmailsResult listEmailsResult = ListingService.ListEmails(userName, id, includeConversationID, sortedBy, sortDirection);
                result = Request.CreateResponse(HttpStatusCode.OK, listEmailsResult);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }
        /// <summary>
        /// Get the list of immediate child emails from a parent CS Folder, paginated
        /// </summary>
        /// <remarks>
        /// Return a List of Content Server Email informations, paginated
        /// </remarks>
        /// <param name="id">The Content Server ID of the parent folder</param>
        /// <param name="includeConversationID">Indicate if you want to retrieve the conversation ID for the emails</param>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <param name="pageNumber">The page number, default is 1</param>
        /// <param name="pageSize">The page size, default is 200</param>
        /// <param name="sortedBy">
        /// Default is OTEmailReceivedDate.
        /// Available options:
        /// OTEmailFrom
        /// OTEmailSentDate
        /// OTEmailTo
        /// OTEmailReceivedDate
        /// OTEmailCC
        /// OTEmailHasAttachments
        /// OTEmailSubject
        /// </param>
        /// <param name="sortDirection">Direction of the sort. Either "asc" or "desc", default is "desc"</param>
        /// <returns></returns>
        [ResponseType(typeof(ListPaginatedEmailsResult))]
        [HttpGet, Route("list/paginated")]
        public async Task<IHttpActionResult> GetEmailsPaginated(long id, bool includeConversationID = false, string userName = "", int pageNumber = 1, int pageSize = 200, string sortedBy = "", string sortDirection = "desc")
        {

            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                var userGroupsResult = CSAccess.CurrentUserGroupsCookie(Request, userName);
                string userGroups = "";
                foreach (var cookie in Request.Headers.GetCookies())
                {
                    var ebidCookie = cookie.Cookies.FirstOrDefault(c => c.Name == "ebid");
                    if (ebidCookie != null)
                    {
                        userGroups = SecureInfo.readSensitiveInfo(ebidCookie.Value.ToString());
                    }
                }

                ListPaginatedEmailsResult listEmailsResult = AGODataAccess.GetPaginatedEmails(id, pageNumber, pageSize, sortedBy, sortDirection, ((!string.IsNullOrEmpty(userGroups)) ? userGroups : userGroupsResult.Item2));//ListingService.ListEmails(userName, id, includeConversationID, pageNumber, pageSize, sortedBy, sortDirection);
                result = Request.CreateResponse(HttpStatusCode.OK, listEmailsResult);

                if (!userGroupsResult.Item1)
                {
                    string usergroups = userGroupsResult.Item2;
                    List<CookieHeaderValue> cookies = new List<CookieHeaderValue>();
                    cookies.Add(new CookieHeaderValue("ebid", usergroups)
                    {
                        Path = "/"
                    });

                    result.Headers.AddCookies(cookies);
                }

            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }
        /// <summary>
        /// Find the emails that also belongs to the same conversation as the specified email
        /// </summary>
        /// <remarks>
        /// Return a List of Content Server Email informations
        /// </remarks>
        /// <param name="emailID">The Content Server ID of the specified email</param>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <param name="sortedBy">
        /// Default is OTEmailReceivedDate.
        /// Available options:
        /// OTEmailFrom
        /// OTEmailSentDate
        /// OTEmailTo
        /// OTEmailReceivedDate
        /// OTEmailCC
        /// OTEmailHasAttachments
        /// OTEmailSubject
        /// </param>
        /// <param name="sortDirection">Direction of the sort. Either "asc" or "desc", default is "desc"</param>
        /// <returns></returns>
        [ResponseType(typeof(ListEmailsResult))]
        [HttpGet, Route("conversation")]
        public async Task<IHttpActionResult> FindConversationEmails(long emailID, string userName = "", string sortedBy = "", string sortDirection = "desc")
        {

            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                ListEmailsResult listEmailsResult = new ListEmailsResult();
                listEmailsResult = ListingService.FindConversationEmails(userName, emailID, sortedBy, sortDirection);
                result = Request.CreateResponse(HttpStatusCode.OK, listEmailsResult);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        /// <summary>
        /// Search for emails from Content Server
        /// </summary>
        /// <remarks>
        /// Searches recursively inside a parent CS Folder
        /// Returns the information about the search, as well as a list of emails that matches the search criteria
        /// </remarks>
        /// <param name="keywords">
        /// The keywords of the search, does not have to be url encoded for now. supports spaces currently.
        /// </param>
        /// <param name="folderNodeID">
        /// Search in this folder id, recursively. Default is 2000 for enterprise
        /// </param>
        /// <param name="firstResultToRetrieve">
        /// The start index of the result, default is 1. For example, 
        /// the total number of results is 500, and you have shown results 1 to 200 (with include count being 200), 
        /// you would want the FirstResultToRetrieve to be 201.
        /// </param>
        /// <param name="numResultsToRetrieve">
        /// The page size, how many results do you want to retrieve. Default is 200
        /// </param>
        /// <param name="sortedBy">
        /// Default is OTEmailReceivedDate.
        /// Available options:
        /// OTEmailFrom
        /// OTEmailSentDate
        /// OTEmailTo
        /// OTEmailReceivedDate
        /// OTEmailCC
        /// OTEmailHasAttachments
        /// OTEmailSubject
        /// </param>
        /// <param name="sortDirection">Direction of the sort. Either "asc" or "desc", default is "desc"</param>
        /// <param name="ReceivedDate_From">Received after this date. YYYYMMDD format. e.g. 20200827    default is blank</param>
        /// <param name="ReceivedDate_To">Received before this date. YYYYMMDD format. e.g. 20200827     default is blank</param>
        /// <param name="FilterByAttachment">Default is UNDEFINED</param>
        /// <param name="EmailSubject">subject contains. default is blank</param>
        /// <param name="EmailFrom">from contains. default is blank</param>
        /// <param name="EmailTo">to contains. default is blank</param>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <returns></returns>
        [ResponseType(typeof(EmailSearchResult))]
        [HttpGet, Route("search")]
        public async Task<IHttpActionResult> SearchForEmails(string keywords = "", long folderNodeID = 2000,
            long? folderNodeID2 = null, long? folderNodeID3 = null, long? folderNodeID4 = null,
            int firstResultToRetrieve = 1, int numResultsToRetrieve = 200, string sortedBy = "", string sortDirection = "desc",
            string ReceivedDate_From = "", string ReceivedDate_To = "", SearchFilterByAttachment FilterByAttachment = SearchFilterByAttachment.UNDEFINED,
            string EmailSubject = "", string EmailFrom = "", string EmailTo = "", string userName = "",
            bool includeConversationID = false, string attachmentFileName = "", string cacheId = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                //search in Archive Folder with nodeID defined from config file
                if (folderNodeID2 == 0000)
                {
                    folderNodeID2 = AGOServices.GetArchiveFolderNodeByPath(userName, folderNodeID);
                }

                EmailSearchResult searchResult = EmailSearchService.SearchForEmails(userName, keywords, folderNodeID, folderNodeID2, folderNodeID3, folderNodeID4,
                    firstResultToRetrieve, numResultsToRetrieve, sortedBy, sortDirection,
                 ReceivedDate_From, ReceivedDate_To, FilterByAttachment,
                 EmailSubject, EmailFrom, EmailTo, includeConversationID, attachmentFileName, cacheId);

                result = Request.CreateResponse(HttpStatusCode.OK, searchResult);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        /// <summary>
        /// Get information about an Email
        /// </summary>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <param name="id">node id</param>
        /// <returns></returns>
        [ResponseType(typeof(EmailInfo))]
        [HttpGet, Route("single")]
        public async Task<IHttpActionResult> GetEmailInfo(long id, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                EmailInfo node = AGOServices.GetEmailInfo(userName, id);
                result = Request.CreateResponse(HttpStatusCode.OK, node);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }
        /// <summary>
        /// Gets the latest version content of an email document.
        /// </summary>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <param name="id">node id of the document</param>
        /// <returns></returns>
        [HttpGet, Route("content")]
        public async Task<IHttpActionResult> GetLatestVersion(long id, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {

                MemoryStream outStream = new MemoryStream();
                EmailContentInfo contentInfo = AGOServices.GetEmailContents(userName, id, outStream);

                byte[] content = outStream.ToArray();
                result = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(content)
                };
                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = contentInfo.FileName,
                    CreationDate = contentInfo.CreationDate,
                    ModificationDate = contentInfo.ModificationDate,
                    Size = contentInfo.FileSize
                };
                result.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }
        /// <summary>
        /// Gets the HTML views of latest version content of an email.
        /// </summary>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <param name="id">node id of the document</param>
        /// <returns></returns>
        [HttpGet, Route("viewer")]
        public async Task<IHttpActionResult> GetHTMLEmailView(long id, string userName = "")
        {

            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {

                MemoryStream outStream = new MemoryStream();
                EmailContentInfo contentInfo = AGOServices.GetEmailContents(userName, id, outStream);

                byte[] content = outStream.ToArray();

                string htmlresult = AGOServices.GetEmailPreviewInHTML(id, userName, content);
                result = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(htmlresult)
                };

                result.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/html");
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }
        [ResponseType(typeof(EmailSearchResult))]
        [HttpGet, Route("searchAttachment")]
        public async Task<IHttpActionResult> SearchForAttachmentFileName(string userName = "", string keywords = "", long folderNodeID = 2000,
            long? folderNodeID2 = null, long? folderNodeID3 = null, long? folderNodeID4 = null,
            int firstResultToRetrieve = 1, int numResultsToRetrieve = 200, string sortedBy = "", string sortDirection = "desc",
            string ReceivedDate_From = "", string ReceivedDate_To = "", string EmailSubject = "", string EmailFrom = "", string EmailTo = "", string attachmentFileName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                long userId = AGOServices.GetCurrentUserID(userName);
                EmailSearchResult searchResult = EmailSearchService.SearchAttachmentFileName(userId, keywords, folderNodeID, folderNodeID2, folderNodeID3, folderNodeID4, firstResultToRetrieve, numResultsToRetrieve, sortedBy, sortDirection, ReceivedDate_From, ReceivedDate_To, EmailSubject, EmailFrom, EmailTo, attachmentFileName);

                result = Request.CreateResponse(HttpStatusCode.OK, searchResult);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        [HttpGet, Route("GetEmailSentDates")]
        public async Task<IHttpActionResult> GetEmailSentDates(string userName, string nodeIds)
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                var emailInfos = AGODataAccess.GetEmailSentDates(nodeIds);
                result = Request.CreateResponse(HttpStatusCode.OK, emailInfos);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }
    }
}