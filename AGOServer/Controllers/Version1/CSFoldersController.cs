using NLog;
using System;
using System.Collections.Generic;
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
    /// Content Server Folders
    /// </summary>
    /// <remarks>
    /// Currently we filter by folder and email folders only
    /// </remarks>
    [RoutePrefix("api/v1/folders")]
    public class CSFoldersController : ApiController
    {
        /// <summary>
        /// Get the list of immediate child folders from a parent CS Folder
        /// </summary>
        /// <remarks>
        /// Note that the childcount also includes non folder items and non email items as well.
        /// Returns a List of Content Server folder informations
        /// </remarks>
        /// <param name="id">The Content Server ID of the parent folder</param>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <returns></returns>
        /// 

        [System.Web.Http.HttpGet, System.Web.Http.Route("GetShowSummarizer")]
        public async Task<IHttpActionResult> GetShowSummarizer()
        {
            bool isShowSummarizer = Properties.Settings.Default.ShowSummarizer;
            var result = Request.CreateResponse(HttpStatusCode.OK, isShowSummarizer);

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }
        [System.Web.Http.HttpGet, System.Web.Http.Route("GetShowSearchListFolderLocation")]
        public async Task<bool> GetShowSearchListFolderLocation()
        {
            bool isShowSummarizer = Properties.Settings.Default.ShowSearchListFolderLocation;
            return isShowSummarizer;
        }

        [ResponseType(typeof(IEnumerable<FolderInfo>))]
        [HttpGet, Route("{id}")]
        public async Task<IHttpActionResult> GetSubFolders(long id, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                List<FolderInfo> folderInfos = new List<FolderInfo>();
                folderInfos = AGOServices.ListSubFolders(userName, id);

                result = Request.CreateResponse(HttpStatusCode.OK, folderInfos);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        /// <summary>
        /// Get the list folders which navigates to the specified item
        /// </summary>
        /// <remarks>
        /// For example, if the item is Enterprise/FolderA/FolderA-1, this will return a list of 3 folderInfo, with the first element being "Enterprise", and the last being "FolderA-1".
        /// If the item is Enterprise/FolderA/FolderA-1/Email01, this will also return the same exact list, because the item is an email. this is useful for navigating the the folder in focus for the user.
        /// It will return a list of folder information, from the Root folder to the item. If the item is a folder, it will stop at that item. If the item is a file, it will stop at the parent folder of the item. Note that the childcount also includes non folder items and non email items as well.
        /// </remarks>
        /// <param name="id">The Content Server ID of the item, can be a file or folder</param>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <returns></returns>
        [ResponseType(typeof(IEnumerable<FolderInfo>))]
        [HttpGet, Route("navigationpath/{id}")]
        public async Task<IHttpActionResult> GetNavigationFolders(long id, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                List<FolderInfo> folderInfos = new List<FolderInfo>();
                folderInfos = AGODataAccess.GetNavigationPath(id);//AGOServices.GetNavigationFolders(userName, id);

                result = Request.CreateResponse(HttpStatusCode.OK, folderInfos);
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
        /// <summary>
        /// Get information about an Folder
        /// </summary>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <param name="id">node id</param>
        /// <returns></returns>
        [ResponseType(typeof(FolderInfo))]
        [HttpGet, Route("single")]
        public async Task<IHttpActionResult> GetEmailInfo(long id, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                FolderInfo node = AGOServices.GetFolderInfo(userName, id);
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
        /// Get node by giving full path
        /// </summary>
        /// <param name="userName">Impersonation username, not required if windows authentication is enabled</param>
        /// <param name="rootId">node id of root folder in path </param>
        /// <param name="path">folder path</param>
        /// <returns></returns>
        [ResponseType(typeof(FolderInfo))]
        [HttpGet, Route("getNodeByPath")]
        public async Task<IHttpActionResult> GetFolderNodeByPath(long rootId, string path, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                FolderInfo folder = AGOServices.GetFolderNodeByPath(userName, rootId, path);
                result = Request.CreateResponse(HttpStatusCode.OK, folder);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        /// <summary>
        /// Search and Get List of Folder which contain similar name
        /// </summary>
        /// <param name="userName">Impersonation username, not required if windows authentication is enabled</param>
        /// <param name="folderName">folder name included in search</param>
        /// <returns></returns>
        [ResponseType(typeof(FolderInfo))]
        [HttpGet, Route("getFolderListByName")]
        public async Task<IHttpActionResult> GetFolderListByName(string folderName, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                List<FolderSearchInfo> folder = AGOServices.GetListFolderByName(userName, folderName);
                result = Request.CreateResponse(HttpStatusCode.OK, folder);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        /// <summary>
        /// Get Default Allowed folder under Enterprise
        /// </summary>
        /// <param name="userName">Impersonation username, not required if windows authentication is enabled</param>
        /// <returns></returns>
        [ResponseType(typeof(FolderInfo))]
        [HttpGet, Route("getDefaultFolder")]
        public async Task<IHttpActionResult> GetDefaultFolder(string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                List<FolderInfo> folder = AGOServices.GetAllowedFolderInfo(userName);
                result = Request.CreateResponse(HttpStatusCode.OK, folder);
                List<CookieHeaderValue> cookies = new List<CookieHeaderValue>();
                cookies.Add(new CookieHeaderValue("default-folder", Properties.Settings.Default.FolderUnderEnterprise)
                {
                    Path = "/"
                });
                result.Headers.AddCookies(cookies);
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
