using AGOServer.Components;
using MimeKit;
using MsgKit;
using MsgReader;
using MsgReader.Outlook;
using Newtonsoft.Json;
using OpenText.ContentServer.API;
using OpenText.Livelink.Service.Core;
using OpenText.Livelink.Service.DocMan;
using OpenText.Livelink.Service.MemberService;
using Org.BouncyCastle.Asn1.Ocsp;
using Org.BouncyCastle.Crypto;
using Rebex;
using Rebex.Mail;
using RtfPipe;
using RtfPipe.Tokens;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.ServiceModel.Syndication;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http.Results;
using WebGrease.Activities;
using static MsgReader.Outlook.Storage;

namespace AGOServer
{
    public static class AGOServices
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        internal static List<FolderInfo> GetNavigationFolders(string userNameToImpersonate, long id)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            List<FolderInfo> folders = new List<FolderInfo>();

            Node currentNode = CSAccess.GetNode(userNameToImpersonate, id);

            bool isFolder = currentNode.IsContainer;
            if (isFolder)
            {
                folders.Add(GetFolderInfoFromNode(currentNode));
            }

            while (currentNode.ParentID > 0)
            {
                Node parentNode = CSAccess.GetNode(userNameToImpersonate, currentNode.ParentID);
                if (parentNode != null)
                {
                    folders.Add(GetFolderInfoFromNode(parentNode));
                }
                currentNode = parentNode;
            }
            folders.Reverse();

            stopwatch.Stop();

            TimeSpan elapsedTime2 = stopwatch.Elapsed;
            logger.Info($"Time taken to get the navigation breadcrumb (ImpersonateUsername: {userNameToImpersonate}/FolderID: #{id}): {elapsedTime2.TotalMilliseconds} ms");

            return folders;
        }

        private static FolderInfo GetFolderInfoFromNode(Node currentNode)
        {
            return new FolderInfo
            {
                ParentNodeID = currentNode.ParentID,
                ChildCount = currentNode.ContainerInfo.ChildCount,
                Name = currentNode.Name,
                NodeID = currentNode.ID,
                NodeType = currentNode.Type
            };
        }

        public static List<FolderInfo> ListSubFolders(string userNameToImpersonate, long folderID)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            List<FolderInfo> folders = new List<FolderInfo>();

            /*   string[] definedFolderFromConfig = Properties.Settings.Default.FolderUnderEnterprise.Split(';');
               long[] definedFolder = definedFolderFromConfig.Select(long.Parse).ToArray();

               // Method ListNodesByPage provide filter of ContainersOnly and IncludeHiddenNodes.
               var result = CSAccess.ListNodesByPage(userNameToImpersonate, folderID, 1, 500, true, null);

               var nodes = result.Nodes;

               if (nodes.Length > 1)
               {
                   Array.Sort(nodes,
                   delegate (Node x, Node y) { return x.Name.CompareTo(y.Name); });
               }
               if (nodes != null)
               {
                   foreach (Node node in nodes)
                   {
                       //if (folderID == 2000)
                       //{
                       //    // If it's under Enterprise then filter only for 'Email' and 'Records' based from config file.
                       //    if (Array.Exists(definedFolder, element => element == node.ID))
                       //    {
                       //        folders.Add(new FolderInfo()
                       //        {
                       //            Name = node.Name,
                       //            NodeID = node.ID,
                       //            ParentNodeID = node.ParentID,
                       //            NodeType = node.Type,
                       //            ChildCount = node.ContainerInfo == null ? 0 : node.ContainerInfo.ChildCount
                       //        });
                       //    }
                       //}
                       //else
                       //{
                       folders.Add(new FolderInfo()
                       {
                           Name = node.Name,
                           NodeID = node.ID,
                           ParentNodeID = node.ParentID,
                           NodeType = node.Type,
                           ChildCount = node.ContainerInfo == null ? 0 : node.ContainerInfo.ChildCount
                       });
                       //}
                   }
               }*/

            //string[] definedFolderFromConfig = Properties.Settings.Default.FolderUnderEnterprise.Split(';');
            // long[] definedFolder = definedFolderFromConfig.Select(long.Parse).ToArray();
            long[] definedFolder = AGODataAccess.GetFolderIDAfterEnterprise(2000).ToArray();
            // Method ListNodesByPage provide filter of ContainersOnly and IncludeHiddenNodes.
            var result = CSAccess.ListNodesByPage(userNameToImpersonate, folderID, 1, 500, true, null);

            var nodes = result.Nodes;

            if (nodes.Length > 1)
            {
                Array.Sort(nodes,
                delegate (Node x, Node y) { return x.Name.CompareTo(y.Name); });
            }
            if (nodes != null)
            {
                foreach (Node node in nodes)
                {
                    if (folderID == 2000)
                    {
                        // If it's under Enterprise then filter only for 'Email' and 'Records' based from config file.
                        if (Array.Exists(definedFolder, element => element == node.ID))
                        {
                            folders.Add(new FolderInfo()
                            {
                                Name = node.Name,
                                NodeID = node.ID,
                                ParentNodeID = node.ParentID,
                                NodeType = node.Type,
                                ChildCount = node.ContainerInfo == null ? 0 : node.ContainerInfo.ChildCount
                            });
                        }
                    }
                    else
                    {
                        folders.Add(new FolderInfo()
                        {
                            Name = node.Name,
                            NodeID = node.ID,
                            ParentNodeID = node.ParentID,
                            NodeType = node.Type,
                            ChildCount = node.ContainerInfo == null ? 0 : node.ContainerInfo.ChildCount
                        });
                    }
                }
            }


            stopwatch.Stop();

            TimeSpan elapsedTime2 = stopwatch.Elapsed;
            logger.Info($"Time taken to get sub folders (ImpersonateUser: {userNameToImpersonate}/Folder: #{folderID}): {elapsedTime2.TotalMilliseconds} ms");
            return folders;
        }

        internal static EmailInfo GetEmailInfo(string userNameToImpersonate, long id)
        {
            Node node = CSAccess.GetNode(userNameToImpersonate, id);
            long fileSize = 0;
            if (node.VersionInfo != null)
            {
                fileSize = (long)node.VersionInfo.FileDataSize;
            }
            EmailInfo emailInfo = new EmailInfo()
            {
                Name = node.Name,
                NodeID = node.ID,
                ParentNodeID = node.ParentID,
                FileSize = fileSize
            };
            AGODataAccess.GetEmailInfoFromCSDB(ref emailInfo);
            emailInfo.ConversationId = AGODataAccess.GetEmailConversationIDFromCSDB(emailInfo.NodeID);
            return emailInfo;
        }
        internal static FolderInfo GetFolderInfo(string userNameToImpersonate, long id)
        {
            Node node = CSAccess.GetNode(userNameToImpersonate, id);

            FolderInfo folderInfo = node != null ? GetFolderInfoFromNode(node) : null;
            return folderInfo;
        }
        public static EmailContentInfo GetEmailContents(string userNameToImpersonate, long id, Stream outStream)
        {
            FileAtts fileAtts = CSAccess.GetNodeLatestVersion(userNameToImpersonate, id, outStream);
            EmailContentInfo contentInfo = new EmailContentInfo
            {
                FileName = fileAtts.FileName,
                CreationDate = fileAtts.CreatedDate,
                ModificationDate = fileAtts.ModifiedDate,
                FileSize = fileAtts.FileSize
            };
            return contentInfo;
        }
        public static long GetCurrentUserID(string userNameToImpersonate)
        {
            long userID = -1;
            try
            {
                User user = CSAccess.GetCurrentUser(userNameToImpersonate);
                if (user != null)
                {
                    userID = user.ID;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ex.StackTrace);
            }
            return userID;
        }

        public static long? GetArchiveFolderNodeByPath(string userName, long folderID)
        {
            long? folderNodeID = null;
            try
            {
                long archiveFolderID = long.Parse(Properties.Settings.Default.ArchiveFolderID);
                string[] folderPath = AGOServices.GetFolderPath(userName, folderID, "Archive");

                Node node = CSAccess.GetFolderNodeByPath(userName, archiveFolderID, folderPath);
                FolderInfo folderInfo = node != null ? GetFolderInfoFromNode(node) : null;
                folderNodeID = folderInfo.NodeID;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ex.StackTrace);
            }

            return folderNodeID;
        }

        internal static string[] GetFolderPath(string userNameToImpersonate, long id, string option)
        {
            List<FolderInfo> folders = new List<FolderInfo>();
            List<String> folder = new List<String>();
            string[] folderPath = null;
            Node currentNode = CSAccess.GetNode(userNameToImpersonate, id);

            if (currentNode != null)
            {
                bool isFolder = currentNode.IsContainer;
                if (isFolder)
                {
                    FolderInfo currentFolder = GetFolderInfoFromNode(currentNode);
                    folders.Add(currentFolder);
                    folder.Add(currentFolder.Name);
                }

                while (currentNode.ParentID > 0)
                {
                    Node parentNode = CSAccess.GetNode(userNameToImpersonate, currentNode.ParentID);
                    if (parentNode != null)
                    {
                        FolderInfo currentFolder = GetFolderInfoFromNode(parentNode);
                        folders.Add(currentFolder);
                        folder.Add(currentFolder.Name);
                    }
                    currentNode = parentNode;
                }
                folders.Reverse();
                folder.Reverse();

                if (option == "Archive")
                {
                    folder.RemoveRange(0, 2);
                }

                folderPath = folder.ToArray();
            }

            return folderPath;
        }

        public static FolderInfo GetFolderNodeByPath(string userName, long id, string path)
        {
            FolderInfo folderInfo = null;
            string[] pathList = path.Split('|');

            Node node = CSAccess.GetFolderNodeByPath(userName, id, pathList);
            folderInfo = node != null ? GetFolderInfoFromNode(node) : null;

            return folderInfo;
        }

        public static List<FolderSearchInfo> GetListFolderByName(string userName, string folderName)
        {

            List<FolderSearchInfo> folderSearchInfos = new List<FolderSearchInfo>();
            List<FolderSearchInfo> listFolder = AGODataAccess.ListFoldersByNameFromCSDB(folderName, GetCurrentUserID(userName));
            List<FolderInfo> allowedFolder = GetAllowedFolderInfo(userName);
            logger.Info($"listFolder {listFolder}");
            if (listFolder.Count > 0)
            {
                foreach (var folder in listFolder)
                {
                    logger.Info($"folder {folder}");
                    FolderSearchInfo folderInfo = new FolderSearchInfo();

                    //get folder path from db wtih db function fn_llpath
                    //string path = AGODataAccess.GetFolderPathFromCSDB(folder.NodeID);
                    //replace string > from db query with /
                    //path = path.Remove(path.Length - 1).Replace(">", " / ");
                    string path = folder.FullPath.TrimEnd('>').Replace(">", " > ");

                    folderInfo.FullPath = path;
                    folderInfo.Name = folder.Name;
                    folderInfo.NodeID = folder.NodeID;
                    folderInfo.ChildCount = folder.ChildCount;
                    folderInfo.ParentNodeID = folder.ParentNodeID;

                    //show only the folder under enterprise
                    if (allowedFolder.Any(e => path.StartsWith("DRMS Workspace > " + e.Name)))
                    {
                        //remove Enterprise text on first path
                        folderInfo.FullPath = folderInfo.FullPath.Replace("DRMS Workspace > ", "");
                        folderSearchInfos.Add(folderInfo);
                    }

                }
            }
            foreach (var folder in folderSearchInfos)
            {
                logger.Info($"folderSearchInfos {folder}");
            }
            return folderSearchInfos;
        }

        public static List<FolderInfo> GetAllowedFolderInfo(string userName)
        {
            HttpResponseMessage result = null;
            List<FolderInfo> allowedFolder = new List<FolderInfo>();
            try
            {
                string[] definedFolderFromConfig = Properties.Settings.Default.FolderUnderEnterprise.Split(';');
                logger.Info($"definedFolderFromConfig  {definedFolderFromConfig.Length}");
                logger.Info($"definedFolderFromConfig length > 1  {definedFolderFromConfig.Length > 0}");
                if (definedFolderFromConfig.Length > 1)
                {
                    long[] definedFolder = definedFolderFromConfig.Select(long.Parse).ToArray();
                    foreach (var folderID in definedFolder)
                    {
                        FolderInfo tempFolderInfo = GetFolderInfo(userName, folderID);
                        if (tempFolderInfo != null)
                        {
                            allowedFolder.Add(tempFolderInfo);
                            logger.Info($"WHAT IS THIS info with config value  {tempFolderInfo}");
                           
                        }
                    }
                    List<CookieHeaderValue> cookies = new List<CookieHeaderValue>();
                    cookies.Add(new CookieHeaderValue("default-folder", Properties.Settings.Default.FolderUnderEnterprise)
                    {
                        Path = "/"
                    });
                    result.Headers.AddCookies(cookies);

                }
                else
                {
                    List<long> getFolderID = AGODataAccess.GetFolderIDAfterEnterprise(2000);
                    foreach (var folderID in getFolderID)
                    {
                        FolderInfo tempFolderInfo = GetFolderInfo(userName, folderID);
                        if (tempFolderInfo != null)
                        {
                            allowedFolder.Add(tempFolderInfo);
                          
                            logger.Info($"WHAT IS THIS info  {tempFolderInfo}");
                        }

                    }
                    List<CookieHeaderValue> cookies = new List<CookieHeaderValue>();
                    cookies.Add(new CookieHeaderValue("default-folder", "")
                    {
                        Path = "/"
                    });
                    result.Headers.AddCookies(cookies);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ex.StackTrace);
            }

            return allowedFolder;
        }

        public static string GetEmailPreviewInHTML(long id, string userName, byte[] content)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            string HtmlResponse = null;
            string previewType = "null";
            //logger.Error("error in processing content byte " + content.ToArray()); 

            try
            {
                Stream storedStream = new MemoryStream(content);
                storedStream.Position = 0;

                //logger.Error("error in processing storedStream " + storedStream.ReadByte().ToString());
                logger.Info($"storedStream Length :  {storedStream.Length}");
                var msg = new MsgReader.Outlook.Storage.Message(storedStream);
                logger.Info($"msg from outlook {msg}");
                logger.Debug($"msg from outlook {msg}");
                var from = msg.Sender.Email;
                var sentDate = msg.SentOn.ToString();
                var receivedDate = msg.ReceivedOn.ToString();
                var recipientsTo = msg.GetEmailRecipients(MsgReader.Outlook.RecipientType.To, false, false);
                var recipientsCc = msg.GetEmailRecipients(MsgReader.Outlook.RecipientType.Cc, false, false);
                var subject = msg.Subject;
                var Header = msg.Headers;
                string results = null;
                var templateHeader = HeaderTemplate(from, sentDate, recipientsTo, receivedDate, recipientsCc, subject);


                if (msg.BodyHtml != null)
                {
                    previewType = "HTML";
                    var bodyHtml = msg.BodyHtml;

                    if (bodyHtml.IndexOf("<body>") < 0)
                    {
                        bodyHtml = $"<body>{bodyHtml}</body>";
                    }

                    // Add <head> tag to the start of the html if does not exist in bodyHtml
                    if (bodyHtml.IndexOf("</head>") < 0)
                    {
                        bodyHtml = "<head><meta charset=\"utf-8\"></head>" + bodyHtml;
                    }


                    results = bodyHtml;

                    Regex charsetPattern = new Regex("charset\\=(.*?)\\\"");
                    foreach (Match match in charsetPattern.Matches(bodyHtml))
                    {
                        if (match.Value.StartsWith("charset"))
                        {
                            results = ReplaceString(results, match.Value, "charset=utf-8\"");
                        }
                    }

                    Regex bodyTagPattern = new Regex(@"<body\s?.*?>");
                    Match bodyTagMatch = bodyTagPattern.Match(bodyHtml);
                    if (bodyTagMatch.Success)
                        results = results.Replace(bodyTagMatch.Value, $"{bodyTagMatch.Value}{templateHeader}");

                    if (msg.Attachments.Count > 0)
                    {
                        for (int attachmentIndex = 0; attachmentIndex < msg.Attachments.Count; attachmentIndex++)
                        {
                            var attachmentObj = msg.Attachments[attachmentIndex];
                            string attachmentName = "";
                            try
                            {
                                if (attachmentObj is Storage.Attachment)
                                {
                                    Storage.Attachment attachment = attachmentObj as Storage.Attachment;
                                    attachmentName = $"[{attachmentIndex}] {attachment.FileName}";
                                    if (!string.IsNullOrEmpty(attachment.ContentId))
                                    {
                                        var data = attachment.Data;
                                        var b64String = Convert.ToBase64String(attachment.Data);
                                        string attachmentMime = (!string.IsNullOrEmpty(attachment.MimeType)) ? attachment.MimeType : GetMimeTypeFromBase64(b64String);

                                        Regex pattern = new Regex(@"(cid:)(" + attachment.ContentId + ")");
                                        if (!string.IsNullOrEmpty(attachment.ContentId))
                                            results = pattern.Replace(results, "data:" + attachmentMime + ";base64," + b64String);
                                    }
                                }
                                else if (attachmentObj is MsgReader.Outlook.Storage.Message)
                                {
                                    Storage.Message msgAttach = attachmentObj as Storage.Message;
                                    attachmentName = $"[{attachmentIndex}] {msgAttach.FileName}";
                                }
                            }
                            catch (Exception ex)
                            {
                                logger.Error($"{previewType} preview/[NodeID: {id} ({userName})] Error processing attachment ({attachmentName}):" + ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                if (attachmentObj is Storage.Attachment)
                                {
                                    Storage.Attachment attachment = attachmentObj as Storage.Attachment;
                                    attachmentName = $"[{attachmentIndex}] {attachment.FileName}";
                                    if (!string.IsNullOrEmpty(attachment.ContentId))
                                    {
                                        var data = attachment.Data;
                                        var b64String = Convert.ToBase64String(attachment.Data);
                                        string attachmentMime = (!string.IsNullOrEmpty(attachment.MimeType)) ? attachment.MimeType : GetMimeTypeFromBase64(b64String);

                                        Regex pattern = new Regex(@"(cid:)(" + attachment.ContentId + ")");
                                        if (!string.IsNullOrEmpty(attachment.ContentId))
                                            results = pattern.Replace(results, "data:" + attachmentMime + ";base64," + b64String);
                                    }
                                }
                                else if (attachmentObj is MsgReader.Outlook.Storage.Message)
                                {
                                    Storage.Message msgAttach = attachmentObj as Storage.Message;
                                    attachmentName = $"[{attachmentIndex}] {msgAttach.FileName}";
                                }
                            }
                            catch (Exception ex)
                            {
                                // Log the error without causing further disruption
                                logger.Error($"msg from outlook {msg}");
                                logger.Error($"{previewType} preview/[NodeID: {id} ({userName})] Error processing attachment ({attachmentName}): {ex.Message}\n{ex.StackTrace}");
                                // Optionally, rethrow the exception to propagate it further
                                // throw;
                            }
                        }
                    }

                    HtmlResponse = $"<!DOCTYPE html>{results}";
                }
                else if (msg.BodyRtf != null)
                {
                    previewType = "RTF";
                    // We cant rely on the attachment index of the msg.Attachments property because it does not match the order of the embedded attachment in the msg.
                    // So we need to sort the attachment list based on the 'RenderingPosition' property nstead for RTF msg.
                    msg.Attachments.Sort((attachmentA, attachmentB) =>
                    {
                        int attachment1Index = -1;
                        int attachment2Index = -1;

                        if (attachmentA is Storage.Attachment)
                        {
                            attachment1Index = (attachmentA as Storage.Attachment).RenderingPosition;
                        }
                        else if (attachmentA is Storage.Message)
                        {
                            attachment1Index = (attachmentA as Storage.Message).RenderingPosition;
                        }
                        else
                        {
                            attachment1Index = -1;
                        }

                        if (attachmentB is Storage.Attachment)
                        {
                            attachment2Index = (attachmentB as Storage.Attachment).RenderingPosition;
                        }
                        else if (attachmentB is Storage.Message)
                        {
                            attachment2Index = (attachmentB as Storage.Message).RenderingPosition;
                        }
                        else
                        {
                            attachment2Index = -1;
                        }

                        return attachment1Index.CompareTo(attachment2Index);
                    });

                    // Remove the illegal parts in the RTF so that it does not break the HTML formatting
                    string rtf = Regex.Replace(msg.BodyRtf, @"\}\\sectd\s?\\(?:ltrsect|rtlsect).*?\{", "}{");

                    string headerStylesheet = @"
<style>
/* Font Definitions */
@font-face
	{font-family:""Cambria Math"";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0cm;
	font-size:11.0pt;
	font-family:""Calibri"",sans-serif;
	mso-fareast-language:EN-US;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:#0563C1;
	text-decoration:underline;}
span.EmailStyle18
	{mso-style-type:personal-reply;
	font-family:""Calibri"",sans-serif;
	color:windowtext;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-size:10.0pt;}
@page WordSection1
	{size:612.0pt 792.0pt;
	margin:72.0pt 72.0pt 72.0pt 72.0pt;}
div.WordSection1
	{page:WordSection1;}
</style>
";
                    var htmlTemplate = "<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\">{0}</head><body>{1}</body></html>";
                    var bodyHtml = templateHeader;
                    var rtfHtml = Rtf.ToHtml(rtf, new RtfHtmlSettings()
                    {
                        AttachmentRenderer = (index, writer) =>
                        {
                            int maxWidthOfDiv = 100;
                            string attachmentName = string.Empty;
                            bool isInline = false;
                            Storage.Attachment attachment = null;
                            Storage.Message msgAttachment = null;

                            if (msg.Attachments[index] is Storage.Attachment)
                            {
                                attachment = (msg.Attachments[index] as Storage.Attachment);
                                attachmentName = attachment.FileName;
                                isInline = attachment.OleAttachment;
                            }
                            else if (msg.Attachments[index] is Storage.Message)
                            {
                                msgAttachment = (msg.Attachments[index] as Storage.Message);
                                attachmentName = msgAttachment.FileName;
                                isInline = false;
                            }

                            FileInfo fileInfo = new FileInfo(attachmentName);

                            if (!isInline)
                            {
                                writer.WriteStartElement("div"); // start of main <div>
                                writer.WriteAttributeString("style", $"width: {maxWidthOfDiv}px; margin: 15px auto;");

                                writer.WriteStartElement("div"); // start of flex <div>
                                writer.WriteAttributeString("style", "display:flex;flex-direction:column; justify-content:center; align-items:center;");

                                writer.WriteStartElement("div"); // <div> for file icon <img>
                                writer.WriteStartElement("img"); // <img> for file icon
                                writer.WriteAttributeString("src", GetFileIconFromExtension(fileInfo.Extension));
                                writer.WriteAttributeString("alt", fileInfo.Name);
                                writer.WriteEndElement(); // end of <img>
                                writer.WriteEndElement(); // end of <div> for file icon <img>

                                writer.WriteStartElement("div"); // <div> for file info
                                writer.WriteAttributeString("style", $"display: flex; flex-direction: column; max-width: {maxWidthOfDiv}px;");
                                writer.WriteStartElement("span"); // <span> for file name
                                writer.WriteAttributeString("style", "box-sizing: border-box; margin: 0; overflow-wrap: break-word; text-align:center; font-size:11pt;font-family:Calibri, sans-serif;");
                                writer.WriteString(fileInfo.Name);
                                writer.WriteEndElement(); // end of <span> for file name
                                writer.WriteEndElement(); // end of file info <div>

                                writer.WriteEndElement(); // end of flex <div>
                                writer.WriteEndElement(); // end of main <div>
                            }
                            else
                            {
                                string inlinePicBase64 = string.Empty;
                                string imgAltTag = string.Empty;

                                writer.WriteStartElement("img"); // <img> for file icon
                                using (MemoryStream outputImg = new MemoryStream())
                                {
                                    using (var source = new MemoryStream(attachment.Data))
                                    {
                                        using (BinaryReader br = new BinaryReader(source))
                                        {
                                            byte[] magicNum = br.ReadBytes(4);
                                            string magicNumStr = BitConverter.ToString(magicNum);

                                            br.BaseStream.Position = 0;
                                            if (magicNumStr == "FF-FF-FF-FF") // When the attachment is an OLE WMF/EMF embedded obj 
                                            {
                                                int length = attachment.Data.Length - 40;
                                                byte[] data = new byte[length];
                                                Buffer.BlockCopy(attachment.Data, 0x28, data, 0, length);
                                                Metafile mf = new Metafile(new MemoryStream(data));

                                                double imageSizeRatio = (double)mf.Width / mf.Height;
                                                int newWidth = (int)Math.Ceiling(70 * imageSizeRatio);
                                                mf.Save(outputImg, ImageFormat.Png);

                                                inlinePicBase64 = Convert.ToBase64String(outputImg.ToArray());

                                                // only for WMF/EMF embedded obj we need to use the width
                                                // attriute to control the attachment preview size
                                                writer.WriteAttributeString("width", "100");
                                            }
                                            else
                                            {
                                                inlinePicBase64 = Convert.ToBase64String(attachment.Data);
                                            }
                                        }
                                    }

                                    writer.WriteAttributeString("src", $"data:{GetMimeTypeFromBase64(inlinePicBase64)};base64,{inlinePicBase64}");
                                    writer.WriteAttributeString("alt", imgAltTag);
                                    writer.WriteEndElement(); // end of <img>
                                }
                            }
                        },
                        ImageUriGetter = picture =>
                        {
                            if (picture.Type is EmfBlip || picture.Type is WmMetafile)
                            {
                                using (var source = new MemoryStream(picture.Bytes))
                                using (var dest = new MemoryStream())
                                {
                                    var bmp = new System.Drawing.Bitmap(source);
                                    bmp.Save(dest, System.Drawing.Imaging.ImageFormat.Png);
                                    return "data:image/png;base64," + Convert.ToBase64String(dest.ToArray());
                                }
                            }
                            return "data:image/png;base64," + Convert.ToBase64String(picture.Bytes);
                        }
                    });

                    bodyHtml += rtfHtml; // Concatenate the header portion with the RTF generated HTML
                    HtmlResponse = string.Format(htmlTemplate, headerStylesheet, bodyHtml); // Fill the blank HTML template with the content from bodyHtml variable

                    Regex charsetPattern = new Regex("charset\\=(.*?)\\\"");
                    foreach (Match match in charsetPattern.Matches(HtmlResponse))
                    {
                        if (match.Value.StartsWith("charset"))
                        {
                            HtmlResponse = ReplaceString(HtmlResponse, match.Value, "charset=utf-8\"");
                        }
                    }
                }
                else
                {
                    previewType = "PLAINTEXT";
                    var bodyHtml = msg.BodyText;
                    HtmlResponse = templateHeader + bodyHtml;
                }
            }
            catch (Exception ex)
            {
                logger.Error($"{previewType} preview/[NodeID: {id} ({userName})] {ex.Message + ex.StackTrace}");
            }

            stopwatch.Stop();

            TimeSpan elapsedTime = stopwatch.Elapsed;
            logger.Info($"Time taken to get the preview (ImpersonateUsername: {userName}/Node: #{id}): {elapsedTime.TotalMilliseconds} ms");

            return HtmlResponse;
        }

        /// <summary>
        /// This method extracts the File Icon of the specified File and retruns the extracted File Icon as Base64 String
        /// </summary>
        /// <param name="filePath">Path of the file to extract its File Icon</param>
        /// <returns>A Base64 Uri of the extracted icon image</returns>
        private static string ExtractBase64FileIcon(string filePath)
        {
            string base64Icon = "data:image/png;base64,";

            if (File.Exists(filePath))
            {
                var icon = System.Drawing.Icon.ExtractAssociatedIcon(filePath);
                using (MemoryStream ms = new MemoryStream())
                {
                    icon.ToBitmap().Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    base64Icon += Convert.ToBase64String(ms.ToArray());
                }
            }

            return base64Icon;
        }

        private static string GetFileIconFromExtension(string extension)
        {
            string base64Uri = "data:image/png;base64,";

            switch (extension)
            {
                case ".docx":
                case ".doc":
                case ".docm":
                case ".dotm":
                case ".dot":
                case ".dotx":
                case ".rtf":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAGeSURBVFhH7ZexSsNQFIbzFn0CF2cfQHyDTuJkwd1BRQs6WdzEB3AUKnZwcemigoJiQBxUHBwELbgUFBEpxCHyH/KXm3LSm9ybIUJ/+Agn3pz/zz03FINKaqe1fdJcX41d2VhbaSSt3IQmURQ5UUoI3wBheOMXwjcArl4hyggAnEOUFQA4hSgzALCGqM2fNmoLZ7EL7fNeyqxzdChmWSSWaWmNi2AGGEf1A3Qu32Noce9+2LzXH8i9qaULqeutO6mxlms0Mw1rgN3jF2mOK+qZ5SupIYbaOniWmmuAZqZhDcC3u376TJlB+903ucddwlo+p5lpWANgm6Gvn1+pu7d9qSGMAvceX7+l5kgAms+2PzLJHQCYM4dgyCAcCcMQGtjIFYBm2HIIs+Yo+DdczWfQfHrzIZNCAXgQqblmOHxzjAYyDyCggY1cAXDaKXOrORrI/EwBmpv1KIUCmJ8eTz7gSCCsMZ+hgY1cAQC32nxT7gy/EBPNTCN3gKJoZhoVDlDiz/E4MgOMCgu1Br5MAkwC/J8Avv+cZoG+iUVVFAR/7jM1CMafiBQAAAAASUVORK5CYII=";
                    break;
                case ".csv":
                case ".xlsx":
                case ".xls":
                case ".xlsm":
                case ".xlsb":
                case ".xltm":
                case ".xlt":
                case ".xltx":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAGHSURBVFhHYxiUoLW5aVNFWfF/cnF5SVEC1CjyAMiQX79+kYWp4ghKHXDy5AnKHEGpA0A0RY6ghgNAmGxHUMsBIEyWI6jpABAm6Aj+KvsE/mqH/+TgRae3oFi2csUysGW4MNRKVIDNYFIwsgPw4aHhAN2eiP8fv3/5DwLLzu2Ai3vPLQCLgQCIjawHm2XYMNEh0LFvAdSq/2AHgcQuP78D5m+9dgRFLQhjswwbJikKHr1/AbYQFAqZazvBbFDIyLX4YKgFGa6wOBWOQXyjI41wTJYDopbWgC0FAZhjKrZNwVAHwjRxAAgfuX8BbDEIgKIAmxoQhllACFPkAFAoYFMDwiDD0fnoIUKyA0DBDQKgeIdFAShxYlNLdQcgZ0WQQ2CJEARguQIZwywghIl2ACirgQBysMNCARQtyGpBGGS4eLQZHIP4Qp7acEySA5BTP8jn2MRBbGQ9VHUAORhmASE8iB1AxeoYH8bpAHQAUojNAErxqANGHTB0HEBp5xQXBpkLtWKwAAYGAKRYm/opOaRMAAAAAElFTkSuQmCC";
                    break;
                case ".ppt":
                case ".pptx":
                case ".pps":
                case ".ppsx":
                case ".pot":
                case ".potx":
                case ".pptm":
                case ".potm":
                case ".ppsm":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAFfSURBVFhH7ZdBSwJBGIb3d9hv8eo96AfYKSyCyCjPRXXK/kBQJ7G6JSEdpUI9eZAudRMSOlhRnpZ06h12ZGaY2dxvplTwg4ddho/vfWYYWDaYytrf270qbOcZlZ2tzWw0ilYYEoYhCS8SrgLNZsNNwlUATycJHwKALOFLAJAkfAqAXyVq6VT2Lr3AKHQrJSXsvFziYTaiSLVMg5MgC8QxOwKm+up/sOfLE6VPoAex+jVn0L5X1hML9J8e2Hurzp+iOqdFpRdg+PAox9jxqiIAho0qXaC9vjRa693e8DXIyL2Ah/6EmwSAOAkngZfqBV+jCACSAEJx5GL3KNM9+DMBvSCk9wJdwAb5BB4PN1hrOaP0yGB4by1jhSwg34E44gRe84uTFfisnE1O4O1gZRSeSCApGI6jFsF4l3f+LwLjYBXw+TmOwyqgFxpNA1yZC8wFZkfA9efUBuZGEdNSQfANfEnn3XL6yUEAAAAASUVORK5CYII=";
                    break;
                case ".png":
                case ".apng":
                case ".jpg":
                case ".jpeg":
                case ".jfif":
                case ".pjpeg":
                case ".pjp":
                case ".bmp":
                case ".gif":
                case ".svg":
                case ".webp":
                case ".ico":
                case ".tif":
                case ".tiff":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAQuSURBVFhH7Zbra5tVHMf7H9jbXohSq+Ad3EBfNLNb1y61yhSHF1Bkir7wtolsbs66FtrmnqZNesm9aZP0os6uvhjzXxBhY7BC3V5MCpomT1vILLhe0vbr73eePE+epFnpGb7TB74kBPJ8Pud3zvmdU/H/U+6ZnJwEZ2JiAuPj40gmk0gkkognEojH4xgbG8Po6ChisVGMxGKIjowgEo0iEokiHI7wb1mv13sg/zr5h+FbW1s7smnMZmk2sb29jampKczPz7PQ/UvwyBmYy+WwwdnQsoF1Levq5xp9qlkXEiywsrKC27d/Rygcvj8JLns5AR3O33cRyGazWF5exs1btxAIhuQleM5ZQMDLCTBcB+8UWFpagqIoSKVSmJ2dhd8fyOZfvbeHF5wukIfvKL8BXCrg8XiK4hscRP7Ve3t4tQsBHb5L+dcKAvw9QyNPZ7RkqBrL8A0MSArQVtMECmADXACNYQlVYNWY1TVaR5vw+iQFeJ/zVtPAT/y0uOc8PbOI/dOLeOGiogv0e31yAtxkVAGuQK4syJirqb/E1rtGn8//qKDhBwWHvlNwNy/Q1++VE+AOx81FK3c5qJYnKQzX0kjg5skM2sYzuoCnv19OgNurqAAL3GMKGPwMlfvAtILrf94R8Bt/3EHrRAbHkhkcjxcEevv65AS4t3MFtEVWCud5fi4/z6bvFTRNKTAT+BUCv07gt8bSeDeWLgh4JAX4YOE9ra1sDfwUgR/8YgYPtLlRd/KSmOcWKvfLVO7XEhm8SeB3CHwimsZH4QVdwN3rkRSgU81YAS73s5cW8dCpGVS2uVD5khOVrXY8/slFvEqjNnfMovnkVRzr/A0fRtL4OLSAUwFNIAeXu1dOgI9UvQKU/TTPD9OIVbADlWYbqo5aKRY89kYQDScuo+nTX4TE8Qtz+NK/gLNDKSHA3dTpdssJ0CmmClD5OXWfT6tgGnWVWQVXtfSgurkL1Uc6UdfqhOn9K7rE2+1z6PSRwN1VIeBwSQoEQyHkSEDraCqYR62BuwW4uqkDNYcvoOZQOx5pdRVJvHd+Dn/nBexOl5xAIBgsEhDlbrGoYDFqAjcR+PC3BP8GtY3nUfviOdSbHQaJawUBh1NOwB8oETCUu1oHt+fBX6P24FnsM32FfQ1nUH/UBtMHP+PIZ7/qAja7Q1LAHxDbh3s5p1BuFVzTSKMW4HOoZbDpDMFP66lvseIgSWgCVrtdTmDY7xcCvIo5arkJbCi3OupisDGP0nRoAhabpMDQ8HCRwD1DAIbsFhbosdrkBAaHdhNY3RNYCwt0W6xyAgODQ6KDaTAZYGmEQI9FToDvcPzHci+UDb+nS1qA7nD8x38rnd09cgJ8h+NrFN9k+DLB5zkfqXyq8cHCvZ3bK3c4bjK8z3mr8WrnBcdzzmXnkTO8o6tbTuA/9FRU/AMd1Tjdu2OmfAAAAABJRU5ErkJggg==";
                    break;
                case ".pdf":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAJKSURBVFhH7ZYxaBRBFIanMKCYHCFRjJEkHkLQHEkjyimIRdKkSYpUCnZia0gXRRBEkAMxIFY2IXAkjSJnI6Q4u0DgIFW4XBVSJNdps7f7tnmZf3Y33u3Nhdm9ISj44Gdvdmb+983Mm2PFXxe/hOinjPjpZwSnlZy/7/WJxdAyWSA5XR84aGz+YMdxkuvokN3lxd80eM6BFxYUWpsFVuB+2dCbJ1Bje4tpsKfhZcROIggA6AzTKBWETQAoMYQpgPfhXVB0QxfY2atqx0QKIGRN9InXYZrOYQpAU1l219cUhLe8pB3Too/vD3A7wjSdwxTAnxhRT7pzi+nJQlt/m+TtgHeYpnMY78BMnt2VgjLFTujGxGUVwP38if0r55luXNb262QVAIWHsSZFGMkqAAqPHs0z3ZtkmpvWjonLKgCNDaizd8ub7A9fZBofUoXpD/eqRDQta+T715Y5VgC8t6+UuTr/2QdM2UtBwpvX/hyFfKqdmcy2zO0KAH88NNofJL07wXR/Sh0DbgJ2wXv+LIBCYinUBj283eKRGkCdtUyMlXlvXgaVrys89KM2nj5WwPH+1AB4572QxvIPB79xBeNjTJQaADuA91i5bmWmSg1gS/8+QC6XUyoUCiftfD7PpVKprU8nKwD1el09o3atVlMQ0bvTZAUACZsBomck7EbznGZZAcBqi8XiSbtSqZztDsTbSF4ul88GoFsZA+AjUmfQjeBpBCA/HL+541erjequ1iiN4AVP+Wm+GqY5PSTECmhtCp6h/f9oCiGOAdsR1317jHm8AAAAAElFTkSuQmCC";
                    break;
                case ".exe":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAEaSURBVFhH7ZTbCoJAEIYlgoggguhZiw5QVBdB14HQ00T0CqUP4AN41puJAVe92F3HRZegHfgQFvH7/1nQMmPmZ+Z8uYJOCm01vJe64PF8cZ+Ftho89DxPC8IAeZ73QpZlJWmattsAfsBavsk0yRsD3Ox7ST3A4uTC/OjC7ODCdO/AZOfAeOvAaPOB4foDg1UVwLZtIUmSqG2AIq9vgNcc5coBKHIWgNec0RhAdAUUOSJrjsRxrLYBihxBMa85QzkARY7ImjOkAURXQJEjKOY1Z0RRpLYBihyRNUe5cgCKHEEprzmjMYDoCqjImiNhGKptgApvA3V57wFkzbUGEMmDIGgfAKH84ShypQBdyn3fFwfQSaE1Y+bvx7K+Vs0alqBeFFIAAAAASUVORK5CYII=";
                    break;
                case ".html":
                case ".htm":
                    base64Uri = "data:image/png;base64,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";
                    break;
                case ".txt":
                case ".xml":
                case ".json":
                case ".config":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAKhSURBVFhH7ZbJjtpAFEX7k/KDiZp5MpjBDGYwMzu+hk160REbFBBIlpDCDC91yzZtoEC8Vu+Skq4wC9c997nqVb38H6rR7/cJ6vV61O12qdPpULvdoVa7Ta1Wi5rNJlmWRY2GRfVGg2r1OlVrNapWa2SaVWpYlm0Yxjd3Ov6A+fF4vNHBr8O1DnQ6nWgwGNB4PAbU5yGQHIb7/Z520M7Tjraets7vRvw62koIACyXS3p//0VmtWrrn4FA2VUAZ3M8PwCwbZsWiwX9fHujcsW0dV3nQeCbA0CaqwBgfja+BZjP5zSbzWgymdBoNKJSqWy7Uz83sOAAMBwOWfIANE27UDafJ3fq5wZWu6yAm/xh+TfbcwXwPBXJf089TUU1FpTN5ZgAYqt5AB/GPnNp6BcgHIC1X+uNWEcH0rNMAOxzbDVVmbkCQEbP8gDQZAAgE/tSX5Repn6QfrOhlVuBdEbnAaDDobmojfHfM3cBhOElwEaW3wPQMhkeANrrV36CVDrNA0BvRwU+Ugq5Kc+plckv03sVSGlMABws2NNnEz8Iw9wDSKY0JoA41VABVUm5QjtPJFM8AByp5wpcJb5NrU7uCa08nkzyACqm6QCIyR8bQ/fNV6u1BIglmADlSoX2AkBVUq4AEI0neAClclkCqBNDTupHyf+4kgCxOA+gWLoGEGaehMFdY8hn7gFEojEmQLEkt4+qpFwBIByN8gCMYlECKBOqdJXaLwCEIkyAgmE8B6AwvBYAguEIDyBfcABUJeUKAIFQmAeQyxdkB1ut1zLlM0nvSQIEQzwA3OHwompCrjDPKxtA3OHw4lfpRyDIA8AdDtco3GRwmcB5jiMVpxoOFvR2tFd0ODQZ7HNsNax2LDh8c5QdyWH+/TXAA/iHxsvLX3ze/jhGI9VzAAAAAElFTkSuQmCC";
                    break;
                case ".zip":
                case ".rar":
                case ".tar":
                case ".7z":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARHSURBVFhH7ZffT5tVGMf5E7z3P/DKuHmnMcbxa3DlhTFeSMaFWYhTHKVF5QLj0LmNOYHVCxTJlpAFSJFZ+4uQjUK3OsuqkhRtaRnQlp8r0hudoX59vud923e1e6HFeWHiSZ6kAfp+Pud7znmfQ9V/eiwsLIzPz89jbm4OwWAQfr8fk5OT4/qvDz+yPzYi+0MjdlU1YDes1a+su1qlb9WC8LGxMQwNDWFgYACDg4NwuVxIztQg8329qvt36rF9py6sP7q8QTjWPt+3souX4L/pQ1dXF2KxmCqLxYKRa1eQjV0E0pel+oEdr5LQH13e4MwJyS11yAMmS+D5ujLYC7fbXRBwOBw480GbAU/3ARkR+K6uUoEGBThI4Iv+DjidzoLA6Ogoes+dMuApCniwXbGArDUBBwl82dcKr9eL5uZmNDU1YWJiAvbzJw14qlcE3P+ewPWrbYhGo4hEImpDhsNhOL6SBPLw1GeaQLBCAe52Ag4SCLrai+ChUAhB52kDnroE3D+MgBwzAvYXsCPksRXD5V0QclsMeJICLmw9fgG7rDMFrEXwQCAgAnIK8vDkpyLwLbZu11YqcFyBHi2gwXnUKPAwfHp6WgRkCfLwpLwPDiOwM2cmYMB51EKedgVvaWnB888+iampKV1Ahz9egWI4jxrXmzMn3Gazwefzyc/eMeCrPSLgxKa8tvVHlzdKBUrhPGoUYOycOeHsAwUBwlcviMA3hxAI/U3gEXAetQdLF3HW+hxef/kpeDwe9flB4pwBXz3/TwVsukApPH/URvob1Ea0vnFUfS6CczNui0CgQoFMqF7FnktQwGcKJ4BQwtPpNEb6RDwPZ6Xke9vXsTH7EpKz1U/rjz94sI8z9lzCKgJeUzh3u5HAEV1Ah1NkcwS/Jy5gc/ZFrM/UhJduvPCEjth/aAKXDQETODeblsCRhxIQ+Irsg/WryG1cQybYiD8iJ7F79zWszVSXd1vSBPpFoF31czO4EhBoUQKEr3yiot8Nn8BvP50AEu/jz5iNKZS3F3iDYey5OAU8pnBuOEKNBOR7hK+c1QRCr2LvFzmW8fek3hWB6goEJPZcXBqLtFMzONe5OAEKCJy1NYHM7eMi0KrgWOyoVKDXEDCBc71LExD48sci8LUk8IoIvK3gWLRi3V+ugNxgGHsuLp1N+rkZXBOoL06A8OWPlEBWNt7ez28pOGLtWCtXgFcoxq4JuEzh3HCEFiVA+HK3EsjcqhYBuSEJHDFLhQISey4unU26mRmcG644Afke4UpgXATqsLfwpoIj2oa16XIF5AbD2AsCJnCu9/DwsLoVsS339MjfEX7vjAg4RKBWBFoUHNHTInCsEgHz2AtHTYr/DRHOjtjdrcPvfagEtm48g70IBWQi0VakyxXgHY6XiHyxmxVKGstGoEar2RrY7XY1c8I7OzvVUeNu53qrktg5c8JTN4/t6Ij/hz6qqv4C1ijZZabKbnoAAAAASUVORK5CYII=";
                    break;
                case ".mp4": // Video File Types
                case ".avi":
                case ".mpeg":
                case ".ogv":
                case ".ts":
                case ".webm":
                case ".mkv":
                case ".wmv":
                case ".flv":
                case ".mov":
                case ".rm":
                case ".rmvb":
                case ".mpg":
                case ".m4v":
                case ".mp3": // Audio File Types
                case ".3gp":
                case ".3g2":
                case ".wav":
                case ".weba":
                case ".oga":
                case ".opus":
                case ".mid":
                case ".midi":
                case ".aac":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATNSURBVFhH7Zd5TJxVFMUxMTY0MY3FxJBg4xZjgwuh2D+IS2LbKHtZotA2EkErDVK2lqUEQqAVK1iFIIHBAlVoTcUIBhJrGloF0pZNCAx7hykDM6wzU2AYZj2++z7oMAiUAfyvNzmZj2TeOb/vvXtfBrtH9ahWllgshkKhwMTEBMbGxjYo4fujo6NobGyEn59fzKKd7dXd3Q2z2YzNlMFgAK0nCH9//1OLlrYVGeh0Om5IIGazCSbTWjLCaLSIanBwEP39/ejq6jL5+vrGMsvHBOcNljXAaqEkI9fy8CUAqXQIcrmc/93R0YHg4GA6jscF9w3UcoC1wlcGCzLwNfT2EomEP6tUKtTX1yMgICCBWW8MggAWFha4wWoAa4UvAfT09KCi4iemCpSXl6Oqqgo1NTVg1s8z7eQh69X6AGuH0zM14VKNj49zmNbWVoIwM+v3mBx5yHq1OsDqZ07BFEqa0hiQc32IPeuh1+utJqm3t9fErD2ZnuUh69VKgKSz3yAxMxtBx8K5Ao+Gw+/Dj+ETdAxe/sE4EZ+CE3FnMM0AXv5+AM0Dcg4giHzM0Gg0BODB5MRD1itrACMLz8GIfAyBR8IR3ibHp00jiOhU4HirHO/7BOGebASfxyTBxCZm78VReNXOYFarY+GCCECtVtsKoOVbSNuckJGN0+lfIyDkE+Q2NiOnqR3J13vw2flaHPIKwPHoJLYLyTAYTXArlcGtTocjdQrmoePTRD5qtco2AK3WAkDhtAOHPwpFbuYthKTfwT7vOPSlO+CAx2FIh2XsCJKhXjDBuUSKkhsdePrafeTViRnAAr9LtgRwKu084lK/hHfgUdz38EGfeyhiL6SgNtoF7xz0RNgX8YiISYSSAbxQdhdadnxf/S2DY86d7QAwID41C7JRBbwCQuC63x0ub7rjddf9cHZx4wBD92QIi4yDWmeCY9kgOqVjmNcZsTvt5gMAdiHZCjDPF9J4xaacQ3RyJjxZx7/08yye+2UOgRn1CAoLR2R0HL4rKMa5nFxo9UY4lQ6gXaLg/bA76U/ezDRJKpXSdgBaSAAxZ87yHaCO33d5Eq6lKiTVqVAmnsGN2624KlYi+y8Jau7OwbGgF4opNX5omYZDTNUigNF2gPl5C8DJ5AxEJabhoKc/kqsuIepaA9okk8iu6YRneR9euTqJPb/OwontjmNuN2Y1GniIxMj48Q8+TVsE0CMqIQ3DI3K8e8gbJy/8jojCatxu6cSLBQPYe2UMDpWz2HlpEvYFw3gmqwWdklE0Dyq4xxYB6G7XI/J0Kr/tqOFIbx/wROVv1XCunsWTl2dgXyiFfW4/dmT1YldSA/qkIzx4SeSjVG4CgEbQcqVabja6XKi73xK1w/6KGvZVSuz4VoYnEhsQf5EazxJO00STpFRObwXgv+GkkfEpFNR1IrS8CRFlt1BSJ9ygS8GCyGfTAJZwIdgagDpckOWNLQDzD7RJAA1faB0uBD88fBsBthq+KQD6v2Bubo69vX4xXLsCgMbLWnRkBL2a9Hr2Y2VqkgA+YNoYwP9RzJp+ET0cICUlpTI/P/9mcXHxPyKRqI1UVCR8ikRFVioqIhWuqsJCQeSTl5fXyqy9mfbwkA3ULqbXmJy3Qa8yvbH4/BTTsrKz+xdCCfEd9FWpEQAAAABJRU5ErkJggg==";
                    break;
                case ".msg":
                case ".eml":
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAADnSURBVFhH7ZSxDYMwEEUZK6KOqCPqiD70oQ996EMftggDMIAHYAAGoLnoW1wCCIhtEGnuS08gTvf9bWx7IpGI5fsH2pNu2K+4kCRXquua2rbdFHjG8WU5gFKKouhMQXCkonhOGrkAL3jCG2PMBuCGLLt/VqNpmoGZDeiFB7zgyd9/BgBVVVEYnnTysnwNaiagB73wgFe/ZhQAYAZpetMNeJqshkmPcQBmaTZ9TFfNOgCY+5+Mzb5xCsCMd7TLyVkVAIzPNN5t7o7VAZg8f2imaktsFsAVCSABJMBsgD3phhWJRP+W570B3m0AMCYxVooAAAAASUVORK5CYII=";
                    break;
                default:
                    base64Uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAKsSURBVFhH7ZbZjtowFIZ5pL5gq2HfAmHf9+2FuOlcTMUNKgikSEhlhzP+kzhkgidwRr1rLf1KcoG/z8Y+duB/U7XxeEzIaDSi4XBIg8GA+v0B9fp96vV61O12qdPpULvdoVa7Tc1WixrNJjUaTarXG9TudIxSqfTN7o7fAL9cLnc5u3P25kzX65UmkwnNZjNIfV0CIwfwdDrRETnKHOkgc7Cee/G0cjAlILDZbOjt7RfVGw1D/4oEpl0l4MDx7iNgGAat12v6+fpK1Vrd0HWdJ4H/HAImXCUAuAO+F1itVrRcLmk+n9N0OqVKpWrYXT/XsOAcARt+N/0usFdA07QPyRUKZHf9XMNqNwUcuM/0728CeF+Ikf9eyCzEbKwpl88zBcRWkwI3sAtuAt2BhCWwc2e3F+voTHqOKYB9jq32EC4gjgDe3fD93hHI6jmeAIqMJYAZuIHNCJnbqOXTC7cEtrZAJqvzBFDhUFzUYHxLuC0ggKrRSwEtm+UJoLyaM+ABmwHM+VbBLQHApUA6k+EJoLZjBu7AiAT7wOXoHQGNKYCDBXvagbhFGHApkEprTAFxqn2cARvqCxfxwC2BEyVTaZ4AjlRnBjxQJVgxchlU00QqxROo1euWgOjcH4x8Dt9ud6ZAPMkUqNZqdBICaqA7/vA/tkAskeQJVKrVBwIW+BHcEYgneALlildAwGQE4FMw4oJLgWgszhQoV8ztI2EPoXbcYLdAJBbjCZTKZVNABVFGAZaBQDjKFCiWSs8JKIDeQCAUifIECkU/gd1TYBkIBMMRnkC+UDQrmIRxgN6YAqEwTwB3OPxQ1SE36OeFLSDucPjh38qPYIgngDscrlG4yeAygfMcRypONRwsqO0or6hwKDLY59hqWO1YcPjPMe0YOeDfX4I8gX+oBQLvEwJiQgPuVJwAAAAASUVORK5CYII=";
                    break;
            }

            return base64Uri;
        }

        private static string GetMimeTypeFromBase64(string base64)
        {
            string mimeType = "";

            Dictionary<string, string> signatures = new Dictionary<string, string>() {
                { "R0lGODdh", "image/gif" },
                { "R0lGODlh", "image/gif" },
                { "iVBORw0KGgo", "image/png" },
                { "/9j/", "image/jpg" },
                { "Qk02L", "image/bmp" },
                { "AAABAA", " image/x-icon" }
            };

            foreach (var key in signatures.Keys)
            {
                if (base64.IndexOf(key) == 0)
                {
                    mimeType = signatures[key];
                    break;
                }
            }

            return mimeType;
        }

        private static string ReplaceString(string Source, string Find, string Replace)
        {
            int Place = Source.IndexOf(Find);

            string result = (Place >= 0) ? Source.Remove(Place, Find.Length).Insert(Place, Replace) : Source;
            return result;
        }

        public static string GetEmailBodyAndAttachments(long id, string userName, byte[] content)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            string HtmlResponse = null;
            string previewType = "null";

            try
            {
                using (MemoryStream storedStream = new MemoryStream(content))
                {
                    storedStream.Position = 0;

                    var message = new Rebex.Mail.MailMessage();
                    message.Load(storedStream);

                    var from = message.From.ToString();
                    var sentDate = message.Date.ToString();
                    var recipientsTo = string.Join(", ", message.To);
                    var recipientsCc = string.Join(", ", message.CC);
                    var subject = message.Subject;
                    var templateHeader = HeaderTemplate(from, sentDate, recipientsTo, sentDate, recipientsCc, subject);

                    StringBuilder attachmentsInfo = new StringBuilder();

                    foreach (var attachment in message.Attachments)
                    {
                        // Handle attachment here as needed
                        string attachmentInfo = $"Saving {attachment.FileName}, ({attachment.MediaType})";
                        attachmentsInfo.AppendLine(attachmentInfo);
                    }

                   /* string attachmentsHtml = $"<p>Attachments:</p><ul>{attachmentsInfo}</ul>";*/

                    if (message.BodyHtml != null)
                    {
                        previewType = "HTML";
                        var bodyHtml = message.BodyHtml;

                        // Process attachments if any
                        // You can adjust this part to handle attachments inline or differently as required

                        // Construct the HTML response
                        HtmlResponse = $"<html><head><meta charset=\"utf-8\"></head><body>{templateHeader}{bodyHtml}</body></html>";
                    }
                    else if (message.BodyText != null)
                    {
                        previewType = "PLAINTEXT";
                        var bodyText = message.BodyText;

                        // Construct the HTML response
                        HtmlResponse = $"<html><head><meta charset=\"utf-8\"></head><body>{templateHeader}{bodyText}</body></html>";
                    }
                    else
                    {
                        // Handle cases where both HTML and plain text bodies are null
                        HtmlResponse = "No content found in the email.";
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
                HtmlResponse = $"Error occurred: {ex.Message}";
            }

            stopwatch.Stop();
            TimeSpan elapsedTime = stopwatch.Elapsed;
            Console.WriteLine($"Time taken to get the preview (ImpersonateUsername: {userName}/Node: #{id}): {elapsedTime.TotalMilliseconds} ms");

            return HtmlResponse;
        }

        private static string HeaderTemplate(string from, string sentDate, string to, string receivedDate, string cc, string subject)
        {
           
            var profileSrc = Properties.Settings.Default.ProfilePicture;
            /*    string pattern = "\"([^\"]*)\"";*/
            string pattern = "^\"?(?<name>[^\"]*)\"?$";
            /* Regex regex = new Regex(pattern);
             Match match = regex.Match(from);
             logger.Info($"match {match.Value}");
             string formatName = match.Value.Trim().Replace(" ", "+");
             logger.Info($"extraformatNamectedName {formatName}");
             string nameWithQuotes = match.Groups[1].Value;

             bool containsQuotes = Regex.IsMatch(nameWithQuotes, pattern);
             string name = "";*/
            /*    string pattern = @"^[""']?(?<name>[A-Za-z\s()]+)[""']?$";*/
            Regex regex = new Regex(pattern);
            Match match = regex.Match(from);
            string name = "";
            if (match.Success)
            {
                 name = match.Groups["name"].Value;
                logger.Info($"Name extracted from '{from}': {name}");
            }
            else
            {
                name = "";
                logger.Info($"No name found in '{from}'");
            }
          /*  if (containsQuotes)
            {
            name = nameWithQuotes.Replace("\"", "");

            }else
            {
                name = nameWithQuotes;
            }
            logger.Info($"from {name}");*/
            string extractedName = ExtractInitials(name);
            var imgSize = (!string.IsNullOrEmpty(cc)) ? 48 : 40;
            DateTime originalDate = DateTime.ParseExact(sentDate, "ddd, d MMM yyyy HH:mm:ss zzz",
                                                    System.Globalization.CultureInfo.InvariantCulture);

            // Convert to the desired format
            string formattedDate = originalDate.ToString("M/d/yyyy h:mm:ss tt");
            var HeaderTemplate = @"
<div class='MsoNormal'>
  <div  class='subject' style='font-size: 14px; margin-top: 11px;'>" + subject + @"</div>
  <hr />
  <div style='display: flex; gap: 1rem; margin-bottom: 1rem;'>
  <svg
          style='border-radius: 999px; height: '" + imgSize + @"px'
          xmlns=""http://www.w3.org/2000/svg""
          width='" + imgSize + @"px'
          height='" + imgSize + @"px'
          viewBox=""0 0 64 64""
          version=""1.1""
        >
          <rect fill=""#ddd"" cx=""32"" width=""64"" height=""64"" cy=""32"" r=""32"" />
          <text
            x=""50%""
            y=""50%""
            alignment-baseline=""middle""
            text-anchor=""middle""
            font-size=""28""
            font-weight=""400""
            dy="".1em""
            dominant-baseline=""middle""
            fill=""#222""
            style='font-size: 28px;'
          >" +
      extractedName +
          "</text><div style='flex-grow: 1;'><div class='from' style='font-size: 12px;'>" + from + @"</div>
        <div style='display: flex; justify-content: space-between; gap: 1rem'>
          <div style='width: 75%;'>
            <div  class='to' style='font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;'>To: " + to + "</div>";
            if (!string.IsNullOrEmpty(cc))
                HeaderTemplate +=
            "<div  class='cc' style='font-size: 11px'>Cc: " + cc + "</div>";

            HeaderTemplate += @"</div>
          <div class='send-date' style='font-size: 11px;'>" + formattedDate + @"</div>
        </div>
    </div>
  </div>
<hr />
</div>
            ";
            return HeaderTemplate;
        }

        public  static string ExtractInitials(string fullName)
        {
            string initials = fullName.Substring(0, 2).ToUpper();
            return initials;
        }
        public static string GetInitials(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "";

            string[] parts = name.Split(new char[] { ' ', '.' }, StringSplitOptions.RemoveEmptyEntries);
            string initials = "";

            foreach (string part in parts)
            {
                initials += char.ToUpper(part[0]);
                if (initials.Length == 2)
                    break; // Break loop after getting first two initials
            }

            return initials;
        }


        public static List<string> GetEmailInfoByMimeKit(byte[] content)
        {
            List<string> emailInfo = new List<string>();

            // Create a memory stream from the provided content
            using (var stream = new MemoryStream(content))
            {
                // Load the stream into a MimeMessage
                var message = MimeMessage.Load(stream);

                // Add email details to the list
                emailInfo.Add($"Subject: {message.Subject}");
                emailInfo.Add($"From: {message.From}");
                emailInfo.Add($"To: {message.To}");
                emailInfo.Add($"Date: {message.Date}");

                // Add email body to the list
                emailInfo.Add("Body:");
                emailInfo.Add(message.TextBody); // or message.HtmlBody if HTML content

                // Handle attachments
                if (message.Attachments.Count() > 0)
                {
                    emailInfo.Add("Attachments:");
                    foreach (var attachment in message.Attachments)
                    {
                        var fileName = attachment.ContentDisposition?.FileName ?? "attachment";
                        emailInfo.Add($"Attachment: {fileName}");
                    }
                }
                else
                {
                    emailInfo.Add("No attachments found.");
                }
            }

            return emailInfo;
        }
    }
}
