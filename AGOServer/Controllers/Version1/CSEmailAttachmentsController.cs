using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using System.Web.Http.Description;
using System.Web.Http.Results;

namespace AGOServer
{
    /// <summary>
    /// Content Server Email Attachments
    /// </summary>
    /// <remarks>
    /// Attachment that reside in CS Emails
    /// </remarks>
    [RoutePrefix("api/v1/attachments")]
    public class CSEmailAttachmentsController : ApiController
    {
        /// <summary>
        /// Get the list of email attachments from an email
        /// </summary>
        /// <remarks>
        /// Returns a List of Content Server Email Attachment Informations. Please call this and wait for the results <b>everytime</b> they click the email's attachment icon.
        /// </remarks>
        /// <param name="id">CS Email node id</param>
        /// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        /// <returns></returns>
        [ResponseType(typeof(IEnumerable<AttachmentInfo>))]
        [HttpGet,Route("")]
        public IHttpActionResult GetEmailAttachments(long id, string userName = "")
        {
            HttpResponseMessage result = null;
            List<AttachmentInfo> attachmentInfos = new List<AttachmentInfo>();
            attachmentInfos = AttachmentServices.ListAttachments(userName, id);

            result = Request.CreateResponse(HttpStatusCode.OK, attachmentInfos);
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }
        /// <summary>
        /// Get the list of embedded attachments from an email
        /// </summary>
        /// <remarks>
        /// Returns a List of Embedded Attachment Informations. 
        /// </remarks>
        /// <param path="path">Attachment temporary file path</param>
        /// <returns></returns>
        [ResponseType(typeof(IEnumerable<int>))]
        [HttpGet, Route("")]
        public IHttpActionResult GetEmbeddedAttachments(string path)
        {
            HttpResponseMessage result = null;
            List<int> attachmentInfos = new List<int>();
            attachmentInfos = AttachmentServices.CheckEmbeddedEmail(path);

            result = Request.CreateResponse(HttpStatusCode.OK, attachmentInfos);
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        /// <summary>
        /// Get the list of email attachments from an email
        /// </summary>
        /// <remarks>
        /// Return the list of attachments as a string to pass to Oscript
        /// </remarks>
        /// <param name="path">temp path</param>
        /// <returns></returns>
        [ResponseType(typeof(string))]
        [HttpPost, Route("listAttachments")]
        public IHttpActionResult GetEmailAttachments(string path)
        {
            HttpResponseMessage result = null;
            string attachmentInfos = AttachmentServices.ListOfAttachments(path);

            result = Request.CreateResponse(HttpStatusCode.OK, attachmentInfos);
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        /// <summary>
        /// Download email attachments from an email
        /// </summary>
        /// <remarks>
        /// Return the list of attachments as a string to pass to Oscript
        /// </remarks>
        /// <param name="emailpath">email path</param>
        /// <param name="tempDir">temp directory</param>
        /// <param name="filename">file name</param>
        /// <returns></returns>
        [ResponseType(typeof(string))]
        [HttpPost, Route("downloadAttachment")]
        public IHttpActionResult DownloadAttachment(string emailpath, string tempDir, string filename)
        {
            HttpResponseMessage result = null;
            string attachmentInfos = AttachmentServices.DownloadAttachment(emailpath,tempDir,filename);

            result = Request.CreateResponse(HttpStatusCode.OK, attachmentInfos);
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        /// <summary>
        /// Download email attachments from an email
        /// </summary>
        /// <remarks>
        /// Return the list of attachments as a string to pass to Oscript
        /// </remarks>
        /// <param name="path">temp path</param>
        /// <returns></returns>
        [ResponseType(typeof(string))]
        [HttpPost, Route("getEmbeddedAttachment")]
        public IHttpActionResult GetEmbeddedAttachment(string path)
        {
            HttpResponseMessage result = null;
            string attachmentInfos = AttachmentServices.CheckEmbededEmail(path);

            result = Request.CreateResponse(HttpStatusCode.OK, attachmentInfos);
            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        ///// <summary>
        ///// Get single attachment
        ///// </summary>
        ///// <remarks>
        ///// Returns Attachment Information. This has to be called <b>everytime</b> the user click on download attachment, 
        ///// or when the user tries to preview the attachment (when the user click on the attachment).
        ///// Please wait until this result is returned before allowing the user to download or view this single attachment
        ///// </remarks>
        ///// <param name="id">Attachment node id</param>
        ///// <param name="userName">Impersonation username, will only be used by the api in testing only, required if windows authentication is disabled</param>
        ///// <returns></returns>
        //[ResponseType(typeof(AttachmentInfo))]
        //[HttpGet, Route("single")]
        //public IHttpActionResult IsEmailAttachmentAvailable(long id, string userName = "")
        //{
        //    HttpResponseMessage result = null;
        //    AttachmentInfo attachmentInfo = AttachmentServices.GetAttachment(userName, id);

        //    result = Request.CreateResponse(HttpStatusCode.OK, attachmentInfo);
        //    ResponseMessageResult responseMessageResult = ResponseMessage(result);
        //    return responseMessageResult;
        //}
    }
}
