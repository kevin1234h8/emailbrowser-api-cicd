using AGOServer.Components;
using CSAccess_Local;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Http.Description;
using System.Web.Http.Results;

namespace AGOServer
{
    [RoutePrefix("api/v2/emails")]
    public class CSEmailsVer2Controller : ApiController
    {
        [ResponseType(typeof(ListEmailsResult))]
        [HttpGet, Route("list")]
        public async Task<IHttpActionResult> GetEmails(long id, string userName = "", int pageNumber = 1, int pageSize = 10, string sortedBy = "", string sortDirection = "desc", bool showAsConversation = false)
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                ListEmailsResultVer2 listEmailsResult = ListingService.ListEmailsVer2(id, userName, pageNumber, pageSize, sortedBy, sortDirection, showAsConversation);
                result = Request.CreateResponse(HttpStatusCode.OK, listEmailsResult);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }
            
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        [ResponseType(typeof(EmailSearchResult))]
        [HttpGet, Route("search")]
        public async Task<IHttpActionResult> SearchForEmails(string keywords = "", long folderNodeID = 2000, 
            long? folderNodeID2 = null, long? folderNodeID3 = null, long? folderNodeID4 = null,
            int firstResultToRetrieve = 1, int numResultsToRetrieve = 200, string sortedBy = "", string sortDirection = "desc",
            string ReceivedDate_From = "", string ReceivedDate_To = "", SearchFilterByAttachment FilterByAttachment = SearchFilterByAttachment.UNDEFINED,
            string EmailSubject = "", string EmailFrom = "", string EmailTo = "", string userName = "",
            bool includeConversationID = false, string attachmentFileName = "", int pageNumber = 1, int pageSize = 10)
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                EmailSearchResultVer2 searchResult = EmailSearchService.SearchForEmailsVer2(userName, keywords, folderNodeID,
                folderNodeID2, folderNodeID3, folderNodeID4,
                firstResultToRetrieve, numResultsToRetrieve, sortedBy, sortDirection,
             ReceivedDate_From, ReceivedDate_To, FilterByAttachment,
             EmailSubject, EmailFrom, EmailTo, includeConversationID, attachmentFileName, pageNumber, pageSize);

                result = Request.CreateResponse(HttpStatusCode.OK, searchResult);
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