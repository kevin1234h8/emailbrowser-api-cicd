using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.ServiceModel.Security;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using System.Web.Http.Results;

namespace AGOServer
{
    /// <summary>
    /// Content Server Docs
    /// </summary>
    /// <remarks>
    /// NodeId 
    /// </remarks>
    [RoutePrefix("api/v1/view")]
    public class CSEmailsViewController : ApiController
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        /// <summary>
        /// Get the document
        /// </summary>
        /// <remarks>
        /// get document by nodeId
        /// </remarks>
        /// <param name="nodeId"> nodeID of document to be downloaded</param>
        /// <returns></returns>
        /// [ResponseType(Content)]
        [HttpGet, Route("email")]
        public async Task<IHttpActionResult> GetEmailFile(long id = 0, string userName = "")
        {
            string htmlresult = null;
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
                htmlresult = AGODataAccess.GetEmailHTMLBody(id);
                if (htmlresult == "")
                {
                    try
                    {
                        MemoryStream outStream = new MemoryStream();
                        EmailContentInfo contentInfo = AGOServices.GetEmailContents(userName, id, outStream);

                        byte[] content = outStream.ToArray();

                        htmlresult = AGOServices.GetEmailBodyAndAttachments(id, userName, content);
                        AGODataAccess.StoreOrUpdateEmailHTMLBody(id, htmlresult);

                        result = Request.CreateResponse(HttpStatusCode.OK, htmlresult);
                        result.Content = new StringContent(htmlresult);
                        result.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                    }
                    catch (Exception ex)
                    {
                        if (userName == null | userName == "")
                        {
                            result = Request.CreateResponse(HttpStatusCode.BadRequest, "userName is required.");
                        }
                        else if (ex.StackTrace.Contains("GetEmailContents"))
                        {
                            result = Request.CreateResponse(HttpStatusCode.NotFound, "Error Processing NodeID.");
                        }
                        else if (ex.StackTrace.Contains("GetEmailPreviewInHTML"))
                        {
                            result = Request.CreateResponse(HttpStatusCode.NoContent, "Error Processing Preview to HTML.");
                        }
                        else
                        {
                            result = Request.CreateResponse(HttpStatusCode.ExpectationFailed, ex.Message.ToString());
                        }
                        logger.Error(ex.StackTrace);
                    }
                }
                else
                {
                    result = Request.CreateResponse(HttpStatusCode.OK, htmlresult);
                    result.Content = new StringContent(htmlresult);
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                }
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        [HttpPost, Route("emailhtml")]
        public async Task<IHttpActionResult> StoreEmailHtmlBody(long id = 0)
        {
            HttpResponseMessage result = null;

            Task.Run(async () =>
            {
                MemoryStream outStream = new MemoryStream();
                EmailContentInfo contentInfo = AGOServices.GetEmailContents(Properties.Settings.Default.DefaultUserName, id, outStream);
                byte[] content = outStream.ToArray();
                var htmlBody = AGOServices.GetEmailPreviewInHTML(id, Properties.Settings.Default.DefaultUserName, content);
                AGODataAccess.StoreOrUpdateEmailHTMLBody(id, htmlBody);
            });

            result = Request.CreateResponse(HttpStatusCode.OK, "Email HTML Body saved successfully.");

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        [HttpGet, Route("emailRebex")]
        public async Task<IHttpActionResult> GetEmailBodyAndAttachmentData(long id = 0, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {

                try
                {
                    MemoryStream outStream = new MemoryStream();
                    EmailContentInfo contentInfo = AGOServices.GetEmailContents(userName, id, outStream);

                    byte[] content = outStream.ToArray();

                    string htmlresult = AGOServices.GetEmailBodyAndAttachments(id, userName, content);
                    /* List<string> htmlresult = AGOServices.GetEmailBodyAndAttachments(id, userName, content);*/
                    /*     AGODataAccess.StoreOrUpdateEmailHTMLBody(id, htmlresult);*/

                    /*   result = Request.CreateResponse(HttpStatusCode.OK, htmlresult);*/
                    /*result.Content = new StringContent(htmlresult);*/
                    /*                result.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");*/
                    return Ok(htmlresult);
                }
                catch (Exception ex)
                {
                    if (userName == null | userName == "")
                    {
                        result = Request.CreateResponse(HttpStatusCode.BadRequest, "userName is required.");
                    }
                    else if (ex.StackTrace.Contains("GetEmailContents"))
                    {
                        result = Request.CreateResponse(HttpStatusCode.NotFound, "Error Processing NodeID.");
                    }
                    else if (ex.StackTrace.Contains("GetEmailPreviewInHTML"))
                    {
                        result = Request.CreateResponse(HttpStatusCode.NoContent, "Error Processing Preview to HTML.");
                    }
                    else
                    {
                        result = Request.CreateResponse(HttpStatusCode.ExpectationFailed, ex.Message.ToString());
                    }
                    logger.Error(ex.StackTrace);
                }

            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.Unauthorized);
            }

            ResponseMessageResult responseMessageResult = ResponseMessage(result);
            return responseMessageResult;
        }

        [HttpGet, Route("email/mimekit")]
        public async Task<IHttpActionResult> GetEmailMimekit(long id = 0, string userName = "")
        {
            HttpResponseMessage result = null;

            if (await CSAccess.ValidateOTSession(Request, userName))
            {
              
                    try
                    {
                        MemoryStream outStream = new MemoryStream();
                        EmailContentInfo contentInfo = AGOServices.GetEmailContents(userName, id, outStream);

                        byte[] content = outStream.ToArray();

                        List<string> htmlresult = AGOServices.GetEmailInfoByMimeKit(content);
                   /*     AGODataAccess.StoreOrUpdateEmailHTMLBody(id, htmlresult);*/

                     /*   result = Request.CreateResponse(HttpStatusCode.OK, htmlresult);*/
                        /*result.Content = new StringContent(htmlresult);*/
        /*                result.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");*/
                        return Ok(htmlresult);
                    }
                    catch (Exception ex)
                    {
                        if (userName == null | userName == "")
                        {
                            result = Request.CreateResponse(HttpStatusCode.BadRequest, "userName is required.");
                        }
                        else if (ex.StackTrace.Contains("GetEmailContents"))
                        {
                            result = Request.CreateResponse(HttpStatusCode.NotFound, "Error Processing NodeID.");
                        }
                        else if (ex.StackTrace.Contains("GetEmailPreviewInHTML"))
                        {
                            result = Request.CreateResponse(HttpStatusCode.NoContent, "Error Processing Preview to HTML.");
                        }
                        else
                        {
                            result = Request.CreateResponse(HttpStatusCode.ExpectationFailed, ex.Message.ToString());
                        }
                        logger.Error(ex.StackTrace);
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
