using NLog;
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
    /// Content Server Users
    /// </summary>
    /// <remarks>
    /// Users that reside in CS
    /// </remarks>
    [RoutePrefix("api/v1/users")]
    public class CSUsersController : ApiController
    {
        /// <summary>
        /// Get the user id of the current logged in user
        /// </summary>
        /// <remarks>
        /// This ID can be used to retrieve the user's personal workspace. -1 if the user is not found. 
        /// Use this ID as what you will use to Get a folder, or to list down the subfolders, note that not all users will have a personal workspace.
        /// </remarks>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <returns></returns>
        [ResponseType(typeof(long))]
        [HttpGet, Route("single")]
        public async Task<IHttpActionResult> GetCurrentUserID(string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                long userID = AGOServices.GetCurrentUserID(userName);
                result = Request.CreateResponse(HttpStatusCode.OK, userID);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        [HttpGet, Route("GetCurrentUserGroups")]
        public async Task<IHttpActionResult> GetCurrentUserGroups(string userName)
        {
            HttpResponseMessage result = null;
            var authRes = await CSAccess.ValidateOTSession(Request, userName);
            if (authRes)
            {
                result = Request.CreateResponse(HttpStatusCode.OK);
                var userGroupsResult = CSAccess.CurrentUserGroupsCookie(Request, userName);

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
    }
}
