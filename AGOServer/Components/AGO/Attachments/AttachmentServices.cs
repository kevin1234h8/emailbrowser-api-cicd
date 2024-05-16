using MsgReader.Mime;
using MsgReader.Outlook;
using Newtonsoft.Json;
using OpenText.Livelink.Service.Core;
using OpenText.Livelink.Service.DocMan;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Web;
using System.Web.Hosting;
using WebGrease.Activities;

namespace AGOServer
{
    public class AttachmentServices
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        //private static PerformanceCounter cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        //private static PerformanceCounter ramCounter = new PerformanceCounter("Memory", "Available MBytes");
        private static List<AttachmentInfo> ExtractAttachments(string userNameToImpersonate, long EmailID)
        {
            List<AttachmentInfo> attachments = new List<AttachmentInfo>();
            byte[] buffer = null;
            using (Stream outStream = new MemoryStream())
            {
                try
                {
                    FileAtts emailAtts = CSAccess.GetNodeLatestVersion(userNameToImpersonate, EmailID, outStream);
                    buffer = (outStream as MemoryStream).ToArray();
                }
                catch (Exception ex)
                {
                    logger.Error("error getting email version " + ex.Message + ex.StackTrace);
                }
            }

            try
            {
                if (buffer != null)
                {
                    using (MemoryStream stream = new MemoryStream(buffer))
                    using (var msg = new MsgReader.Outlook.Storage.Message(stream))
                    {
                        try
                        {
                            if (msg != null && msg.Attachments != null && msg.Attachments.Count > 0)
                            {
                                long emailAttachmentCacheFolderID = -1;
                                ExtractedEmailInfo extractedEmailInfo = AGODataAccess.GetExtractedEmailInfo(EmailID);
                                if (extractedEmailInfo == null)
                                {
                                    extractedEmailInfo = PrepareEmailForExtraction(userNameToImpersonate, EmailID);
                                    //if (extractedEmailInfo != null)
                                    //{
                                    //    emailAttachmentCacheFolderID = extractedEmailInfo.CSCachingFolderNodeID;
                                    //}
                                }
                                //else
                                //{
                                //    emailAttachmentCacheFolderID = extractedEmailInfo.CSCachingFolderNodeID;
                                //}

                                // if (emailAttachmentCacheFolderID > 0)
                                    NodeRights emailNodeRights = CSAccess.GetNodeRights(EmailID);
                                    foreach (var attach in msg.Attachments)
                                    {
                                        try
                                        {
                                            if (attach.GetType() == typeof(MsgReader.Outlook.Storage.Attachment))
                                            {
                                                MsgReader.Outlook.Storage.Attachment attachment = (MsgReader.Outlook.Storage.Attachment)attach;
                                                string attachmentFilename = attachment.FileName;
                                                //Node attachmentNode = CSAccess.GetNodeByNameAsAdmin(emailAttachmentCacheFolderID, attachmentFilename);
                                                //if(attachmentNode == null)
                                                //{
                                                //    attachmentNode = CSAccess.CreateNodeAsAdmin(parentID: emailAttachmentCacheFolderID, attachmentFilename, "Document", attachmentFilename, attachment.Data);
                                                //}
                                                //if (attachmentNode != null && attachmentNode.ID > 0)
                                                //{
                                                //    CSAccess.SetNodeRights(attachmentNode.ID, emailNodeRights);
                                                //}
                                                byte[] attachmentFileHash = GetFileHash(attachment.Data);
                                                int fileExtPosition = attachment.FileName.LastIndexOf('.');
                                                AttachmentInfo attachmentInfo = new AttachmentInfo
                                                {
                                                    CSEmailID = EmailID,
                                                    FileHash = attachmentFileHash,
                                                    CSID = 0, //attachmentNode.ID,
                                                    FileName = attachmentFilename,
                                                    FileSize = attachment.Data.LongLength,
                                                    FileType = (attachment.MimeType != null) ? attachment.MimeType : attachment.FileName.Substring(fileExtPosition, attachment.FileName.Length - fileExtPosition), //attachmentNode.Type,
                                                    Deleted = false
                                                };
                                                AGODataAccess.RegisterEmailAttachment(attachmentInfo);
                                                attachments.Add(attachmentInfo);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            logger.Error("error in processing attachment " + ex.Message + ex.StackTrace);
                                        }
                                    }
                                }
                                else
                                {
                                    logger.Warn("failed to ensure emailAttachmentCacheFolderID exist, it's -1");
                                }
                        }
                        catch (Exception ex)
                        {
                            logger.Error("error in processing attachments " + ex.Message + ex.StackTrace);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("error in extracting or processing attachments" + ex.Message + ex.StackTrace);
            }

            try
            {
                AGODataAccess.UpdateEmailAsExtracted(EmailID);
            }
            catch (Exception ex)
            {
                logger.Error("error in updating email as extracted" + ex.Message + ex.StackTrace);
            }
            return attachments;
        }

        private static ExtractedEmailInfo PrepareEmailForExtraction(string userNameToImpersonate, long emailID)
        {
            //create caching folder(no perm), and insert into EmailBrowser_ExtractedEmails
            ExtractedEmailInfo extractedEmailInfo = null;
            string nameOfAttachmentFolder = emailID.ToString();

            long attachmentSystemRootFolderID = Properties.Settings.Default.AttachmentSystemRootFolderID;
            try
            {
                //Node attachmentFolder = CSAccess.GetNodeByNameAsAdmin(attachmentSystemRootFolderID, nameOfAttachmentFolder);
                //if(attachmentFolder == null)
                //{
                //    attachmentFolder = CSAccess.CreateFolderAsAdmin(attachmentSystemRootFolderID, nameOfAttachmentFolder, "", null);
                //}
                Node emailNode = CSAccess.GetNode(userNameToImpersonate, emailID);
                extractedEmailInfo = new ExtractedEmailInfo
                {
                    CSCachingFolderNodeID = 0, //attachmentFolder.ID,
                    CSID = emailID,
                    IsExtractingNow = true,
                    CSModifyDate = emailNode.ModifyDate.GetValueOrDefault(),
                    CSName = emailNode.Name,
                    CSParentID = emailNode.ParentID
                };
                AGODataAccess.RegisterEmailForExtraction(extractedEmailInfo);
            }
            catch (Exception ex)
            {
                logger.Error("error in PrepareEmailForExtraction email for extraction" + ex.Message + ex.StackTrace);
            }
            return extractedEmailInfo;
        }


        private static byte[] GetFileHash(byte[] bytes)
        {
            using (var md5 = MD5.Create())
            {
                return md5.ComputeHash(bytes);
            }
        }

        public static List<int> CheckEmbeddedEmail(String path)
        {
            List<int> listAttachment = new List<int>();
            try
            {
                //using (var msg = new MsgReader.Outlook.Storage.Message(path))
                //{

                //    var attachments = msg.Attachments;
                //    int count = 0;
                //    foreach (var attachment in attachments)
                //    {
                //        MsgReader.Outlook.Storage.Attachment mObj;
                //        mObj = (MsgReader.Outlook.Storage.Attachment)attachment;
                //        if (mObj.IsInline)
                //        {
                //            listAttachment.Add(count);
                //        }
                //        count++;
                //    }
                //}
            }
            catch (Exception ex)
            {
                logger.Error("Error info:" + ex.Message);
            }
            return listAttachment;
        }

        public static List<AttachmentInfo> ListAttachments(string userNameToImpersonate, long EmailID)
        {
            List<AttachmentInfo> attachments = new List<AttachmentInfo>();

            bool wasExtracted = AGODataAccess.WasAttachmentsExtracted(EmailID);

            if (wasExtracted)
            {
                attachments = AGODataAccess.ListAttachmentsFromCSDB(EmailID);
            }
            else
            {
                attachments = ExtractAttachments(userNameToImpersonate, EmailID);
            }
            if (attachments != null)
            {
                for (int i = 0; i < attachments.Count; i++)
                {
                    attachments[i].No = i;
                }
            }
            return attachments;
        }

        internal static AttachmentInfo GetAttachment(string userName, long id)
        {
            throw new NotImplementedException();
        }

        public static string CheckEmbededEmail(String path)
        {
            string listAttachment = null;
            try
            {
                logger.Info($"Checking for Embedded Attachments within \"{path}\"");
                using (var msg = new MsgReader.Outlook.Storage.Message(path))
                {

                    var attachments = msg.Attachments;
                    int count = 0;
                    foreach (var attachment in attachments)
                    {
                        MsgReader.Outlook.Storage.Attachment aObj;

                        //if (attachment is Storage.Attachment)
                        //{
                        //    aObj = (MsgReader.Outlook.Storage.Attachment)attachment;
                        //    if (!string.IsNullOrEmpty(aObj.ContentId) || aObj.OleAttachment)
                        //    {
                        //        listAttachment = listAttachment + "|" + count.ToString();
                        //    }
                        //}

                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error Getting Embedded Attachment Info for path ({path}):" + ex.Message + $"\n{ex.StackTrace}");
            }
            logger.Info($"Finish checking for embedded attachments within \"{path}\": {listAttachment}");
            return listAttachment;
        }
        public static string ListOfAttachments(string emailPath)
        {
            string retrive = null;
            string path = @"";
            path = Path.Combine(path, emailPath);
            string info = null;
            string fileName = "";
            logger.Info("EMAIL PATH:" + path);
            try
            {
                using (var msg = new MsgReader.Outlook.Storage.Message(path))
                {
                    logger.Info($"Successfully loaded the MSG \"{emailPath}\"");

                    var attachments = msg.Attachments;
                    string outputPath = HostingEnvironment.MapPath("~/MsgContent/");
                    int count = 0;
                    foreach (var attachment in attachments)
                    {
                        MsgReader.Outlook.Storage.Message mObj;
                        MsgReader.Outlook.Storage.Attachment aObj;
                        if (attachment.GetType() == typeof(MsgReader.Outlook.Storage.Message))
                        {
                            mObj = (MsgReader.Outlook.Storage.Message)attachment;
                            logger.Info("MSG Attachment Name:" + mObj.FileName);

                            //the msg obj content and attachment can be read directly just like the parent msg
                            string outDirPath = Path.Combine(outputPath, Path.GetFileName(path));
                            //DirectoryInfo directoryInfo = Directory.CreateDirectory(outDirPath);
                            using (MemoryStream ms = new MemoryStream())
                            {
                                //string savePath = Path.Combine(directoryInfo.FullName, mObj.FileName);
                                mObj.Save(ms);
                                //logger.Info("Saved extracted MSG attachment to: " + savePath);

                                int fileSize = (int)(ms.Length);
                                fileName = Uri.EscapeDataString(mObj.FileName);
                                retrive = count + "," + fileName + "," + fileSize;
                                info = info + "|" + retrive;

                                //File.Delete(savePath);
                                //System.IO.Directory.Delete(directoryInfo.FullName, true);
                            }
                        }
                        else
                        {
                            aObj = (MsgReader.Outlook.Storage.Attachment)attachment;
                            logger.Info("Attachment Name:" + aObj.FileName + "\tIs Inline: " + aObj.IsInline + "\tIs Ole: " + aObj.OleAttachment + "\tContent ID: " + aObj.ContentId);
                            string outDirPath = Path.Combine(outputPath, Path.GetFileName(path));
                            //DirectoryInfo directoryInfo = Directory.CreateDirectory(outDirPath);
                            
                            int fileSize = aObj.Data.Length;
                            fileName = Uri.EscapeDataString(aObj.FileName);
                            retrive = count + "," + fileName + "," + fileSize;
                            info = info + "|" + retrive;
                        }

                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error info: {ex.Message}\r\n{ex.StackTrace}",ex);
            }
            return info;
        }

        public static string DownloadAttachment(string emailpath, string tempDir, string filename)
        {
            string retrive = null;
            string info = null;
            logger.Info($"Downloading attachment: {filename} stored at \"{emailpath}\" in {tempDir}");
            using (var msg = new MsgReader.Outlook.Storage.Message(emailpath))
            {
                try
                {
                    LogMachineResources($"[Begin extracting {emailpath} attachment ({filename}) for download]");
                    var attachments = msg.Attachments;
                    FileInfo fileInfo = new FileInfo(emailpath);
                    string nodeId = fileInfo.Name.Split('_')[3];

                    int count = 0;
                    foreach (var attachment in attachments)
                    {
                        MsgReader.Outlook.Storage.Message mObj;
                        MsgReader.Outlook.Storage.Attachment aObj;
                        string fileName = "";
                        string savePath = "";

                        //the msg obj content and attachment can be read directly just like the parent msg

                        if (attachment.GetType() == typeof(MsgReader.Outlook.Storage.Message))
                        {
                            mObj = (MsgReader.Outlook.Storage.Message)attachment;
                            fileName = mObj.FileName;
                            savePath = Path.Combine(tempDir, fileName);

                            if (filename.ToLower().Contains(fileName.ToLower()))
                            {

                                if (AGODataAccess.IsFromMigratedFolder(nodeId)/* && !mObj.IsDraft*/)
                                {
                                    logger.Info($"Attempting to re create attachment");
                                    string tempAttachmentPath = Path.Combine(tempDir, fileName) + "_temp";
                                    var newEmailAttachment = RecreateEmail(mObj, tempAttachmentPath);
                                    newEmailAttachment.Save(savePath);

                                    //var tempDirInfo = new DirectoryInfo(tempAttachmentPath);
                                    //tempDirInfo.Delete(true);
                                }
                                else
                                {
                                    logger.Info($"Attempting to extract msg attachment");
                                    mObj.Save(savePath);
                                }

                                info = savePath;
                                break;
                            }
                        }
                        else
                        {
                            aObj = (MsgReader.Outlook.Storage.Attachment)attachment;
                            fileName = aObj.FileName;
                            savePath = Path.Combine(tempDir, fileName);

                            if (filename.ToLower().Contains(fileName.ToLower()))
                            {
                                logger.Info("downloading normal attachment");
                                if (aObj.OleAttachment)
                                {
                                    using (var source = new MemoryStream(aObj.Data))
                                    {
                                        using (BinaryReader br = new BinaryReader(source))
                                        {
                                            byte[] magicNum = br.ReadBytes(4);
                                            string magicNumStr = BitConverter.ToString(magicNum);
                                            logger.Info($"Ole Attachment \"{fileName}\" Magic Num: {magicNumStr}");

                                            br.BaseStream.Position = 0;
                                            if (magicNumStr == "FF-FF-FF-FF") // When the attachment is an OLE WMF/EMF embedded obj 
                                            {
                                                int length = aObj.Data.Length - 40;
                                                byte[] data = new byte[length];
                                                Buffer.BlockCopy(aObj.Data, 0x28, data, 0, length);
                                                System.Drawing.Imaging.Metafile mf = new System.Drawing.Imaging.Metafile(new MemoryStream(data));

                                                double imageSizeRatio = (double)mf.Width / mf.Height;
                                                int newWidth = (int)Math.Ceiling(70 * imageSizeRatio);
                                                logger.Info($"Saving WMF file to: {savePath}");
                                                mf.Save(savePath, System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            else
                                            {
                                                logger.Info($"Saving Ole file to: {savePath}");
                                                System.IO.File.WriteAllBytes(savePath, source.ToArray());
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (AGODataAccess.IsFromMigratedFolder(nodeId) && filename.EndsWith(".msg"))
                                    {
                                        logger.Info($"Attempting to re create attachment");
                                        string tempAttachmentPath = Path.Combine(tempDir, fileName) + "_temp";

                                        using (MemoryStream ms = new MemoryStream(aObj.Data))
                                        {
                                            var newEmailAttachment = RecreateEmail(new Storage.Message(ms), tempAttachmentPath);
                                            newEmailAttachment.Save(savePath);

                                            //var tempDirInfo = new DirectoryInfo(tempAttachmentPath);
                                            //tempDirInfo.Delete(true);
                                        }
                                    }
                                    else
                                    {
                                        logger.Info("normal attachment download 2");
                                        System.IO.File.WriteAllBytes(savePath, aObj.Data);
                                    }
                                }
                                info = savePath;
                                break;
                            }
                        }

                        //int fileSize = System.IO.File.ReadAllBytes(savePath).Length;
                        count++;
                    }
                }
                catch (Exception ex)
                {
                    logger.Error($"Error downloading \"{filename}\": {ex.Message}\r\n{ex.StackTrace}", ex);
                }
                // etc...
            }

            LogMachineResources($"[Finish extracting {emailpath} attachment ({filename}) for download]");
            logger.Info($"Result of download: {info}");
            return info;
        }

        private static MsgKit.Email RecreateEmail(MsgReader.Outlook.Storage.Message email, string emailTempFolder = "")
        {
            if (!Directory.Exists(emailTempFolder))
                Directory.CreateDirectory(emailTempFolder);

            MsgKit.Email newEmail = null;

            try
            {
                logger.Info($"Attempting to re-create email: {email.FileName} [Attachments: {email.Attachments.Count} ({email.GetAttachmentNames()}), IsDraft: {email.IsDraft}, HasRtfBody: {!string.IsNullOrEmpty(email.BodyRtf)}, HasHtmlBody: {!string.IsNullOrEmpty(email.BodyHtml)}, HasTextBody: {!string.IsNullOrEmpty(email.BodyText)}]");
                newEmail = new MsgKit.Email(new MsgKit.Sender(email.Sender.Email, email.Sender.DisplayName), email.Subject, email.IsDraft, false);
                newEmail.SentOn = email.SentOn;
                newEmail.ReceivedOn = email.ReceivedOn;
                if (email.Importance != null)
                    newEmail.Importance = (MsgKit.Enums.MessageImportance)email.Importance;
                if (!string.IsNullOrEmpty(email.Id))
                    newEmail.InternetMessageId = email.Id;
                if (!string.IsNullOrEmpty(email.TransportMessageHeaders))
                    newEmail.TransportMessageHeadersText = email.TransportMessageHeaders;
                if (!string.IsNullOrEmpty(email.BodyHtml))
                    newEmail.BodyHtml = email.BodyHtml;
                if (!string.IsNullOrEmpty(email.BodyRtf))
                {
                    newEmail.BodyRtf = email.BodyRtf;
                    newEmail.BodyRtfCompressed = true;
                }
                if (!string.IsNullOrEmpty(email.BodyText))
                    newEmail.BodyText = email.BodyText;

                foreach (var recipient in email.Recipients)
                {
                    if (recipient.Type == MsgReader.Outlook.RecipientType.To)
                        newEmail.Recipients.AddTo(recipient.Email, recipient.DisplayName);
                    else if (recipient.Type == MsgReader.Outlook.RecipientType.Cc)
                        newEmail.Recipients.AddCc(recipient.Email, recipient.DisplayName);
                    else if (recipient.Type == MsgReader.Outlook.RecipientType.Bcc)
                        newEmail.Recipients.AddBcc(recipient.Email, recipient.DisplayName);
                }

                int i = 0;
                foreach (var attachment in email.Attachments)
                {
                    if (attachment == null)
                        continue;

                    bool isInline = false;
                    string fileName = string.Empty;
                    long fileSize = 0;
                    string contentId = string.Empty;
                    int renderingPosition = -1;

                    MemoryStream ms = new MemoryStream();
                    MsgReader.Outlook.Storage.Attachment aObj;
                    MsgReader.Outlook.Storage.Message mObj;

                    if (attachment is MsgReader.Outlook.Storage.Attachment)
                    {
                        aObj = attachment as MsgReader.Outlook.Storage.Attachment;
                        fileName = Path.Combine(emailTempFolder, aObj.FileName);
                        fileSize = aObj.Data.Length;
                        isInline = aObj.IsInline;
                        contentId = aObj.ContentId;
                        renderingPosition = aObj.RenderingPosition;
                        File.WriteAllBytes(fileName, aObj.Data);
                    }
                    else if (attachment is MsgReader.Outlook.Storage.Message)
                    {
                        mObj = attachment as MsgReader.Outlook.Storage.Message;

                        string embeddedMsgTempFolder = Path.Combine(emailTempFolder, $"attch_{i}");
                        var newEmailAttachment = RecreateEmail(mObj, embeddedMsgTempFolder);
                        fileName = Path.Combine(embeddedMsgTempFolder, mObj.FileName);
                        renderingPosition = mObj.RenderingPosition;
                        newEmailAttachment.Save(fileName);
                    }

                    newEmail.Attachments.Add(fileName, renderingPosition, isInline, contentId);
                    i++;
                }
            }
            catch (Exception ex)
            {
                logger.Error("Error Re-Creating Email: " + ex.ToString() + "\r\n" + ex.StackTrace.ToString());
            }

            return newEmail;
        }

        private static void LogMachineResources(string action = "")
        {
            try
            {
                //float cpuUsage = cpuCounter.NextValue();
                //float ramAvailable = ramCounter.NextValue();

                //logger.Info($"{((!string.IsNullOrEmpty(action)) ? $"{action} " : "")}CPU Usage: {cpuUsage}% & RAM Available: {ramAvailable} MB");
            }
            catch (Exception ex)
            {
                logger.Error($"Error getting the machine resources usage: {ex.Message}\r\n{ex.StackTrace}");
            }
        }
    }
}