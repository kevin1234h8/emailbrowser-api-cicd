using AGOServer.Components;
using OpenText.Livelink.Service.Core;
using OpenText.Livelink.Service.DocMan;
using OpenText.Livelink.Service.MemberService;
using System;
using System;
using System.Collections.Generic;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class ListingService
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        public static ListPaginatedEmailsResult ListEmails(string userNameToImpersonate, long folderID, bool includeConversationID, int pageNumber, int pageSize, string paramSortedBy, string sortDirection)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            ListPaginatedEmailsResult result = new ListPaginatedEmailsResult();
            List<EmailInfo> emailInfos = new List<EmailInfo>();
            string sortByColumn = GetSortByColumn(paramSortedBy);

            Stopwatch stopwatch1 = new Stopwatch();
            stopwatch1.Start();
            Rest_ListNodeResult listNodeResult = CSAccess.ListNodesByPage_Rest(userNameToImpersonate, folderID, sortByColumn, sortDirection, pageSize, pageNumber);
            stopwatch1.Stop();
            TimeSpan elapsedTime = stopwatch1.Elapsed;

            logger.Info($"Time taken to get from CS REST API (Impersonate: {userNameToImpersonate}/Folder: #{folderID}/IncludConvID: {includeConversationID}/PageNo: {pageNumber}/PageSize: {pageSize}/SortBy: {paramSortedBy} ({sortDirection})): {elapsedTime.TotalMilliseconds} ms");

            if (listNodeResult != null)
            {
                result.NumberOfPages = listNodeResult.TotalPage;
                result.PageNumber = pageNumber;
                result.PageSize = pageSize;
                result.TotalCount = listNodeResult.TotalCount;
                if (listNodeResult.Nodes != null)
                {
                    foreach (Rest_Node node in listNodeResult.Nodes)
                    {
                        AddNodeToEmailInfos(ref emailInfos, node);
                    }
                }
            }
            RetrieveEmailInfos_AdditionalInfo(includeConversationID, ref emailInfos);
            result.EmailInfos = emailInfos;
            result.SortedBy = paramSortedBy;//for feedback
            result.SortDirection = sortDirection;//for feedback

            // Sort the emails
            //result.EmailInfos = sortEmailInfos(paramSortedBy, sortDirection, result.EmailInfos);


            stopwatch.Stop();

            TimeSpan elapsedTime2 = stopwatch.Elapsed;
            logger.Info($"Time taken to get paginated result (Impersonate: {userNameToImpersonate}/Folder: #{folderID}/IncludConvID: {includeConversationID}/PageNo: {pageNumber}/PageSize: {pageSize}/SortBy: {paramSortedBy} ({sortDirection})): {elapsedTime2.TotalMilliseconds} ms");
            return result;
        }

        /// <summary>
        /// List out mail without pagination
        /// </summary>
        /// <param name="userNameToImpersonate"></param>
        /// <param name="folderID"></param>
        /// <param name="includeConversationID"></param>
        /// <param name="paramSortedBy"></param>
        /// <param name="sortDirection"></param>
        /// <returns></returns>
        public static ListEmailsResult ListEmails(string userNameToImpersonate, long folderID, bool includeConversationID, string paramSortedBy, string sortDirection)
        {
            ListEmailsResult result = new ListEmailsResult();
            List<EmailInfo> emailInfos = new List<EmailInfo>();
            int maxCount = Properties.Settings.Default.FlatEmailListing_MaxCount;
            string sortByColumn = GetSortByColumn(paramSortedBy);

            int page = 1;
            Rest_ListNodeResult listNodeResult = CSAccess.ListNodesByPage_Rest(userNameToImpersonate, folderID, sortByColumn, sortDirection, maxCount, page);

            if (listNodeResult != null && listNodeResult.Nodes != null)
            {
                //a 200 limit should get 125, for 125 count
                //a 200 limit should get 200, for 201 count
                //a 300 limit should get 201, for 201 count

                bool hasNextPage = listNodeResult.TotalPage > page;
                bool willHitLimit = listNodeResult.TotalCount >= maxCount;

                if (hasNextPage)
                {
                    int target;
                    if (willHitLimit)
                    {
                        //make sure limit is reach, because we will hit the limit
                        target = maxCount;
                    }
                    else
                    {
                        target = listNodeResult.TotalCount;
                        //we won't hit the limit, so make sure that RangeMax hits the TotalCount
                    }

                    bool isTargetReach = listNodeResult.Nodes.Count >= target;
                    while (isTargetReach == false)
                    {
                        int newLimit = target - listNodeResult.Nodes.Count;
                        page++;
                        Rest_ListNodeResult nextListNodeResult =
                            CSAccess.ListNodesByPage_Rest(userNameToImpersonate, folderID, sortByColumn, sortDirection, newLimit, page);
                        if (nextListNodeResult != null && nextListNodeResult.Nodes != null)
                        {
                            listNodeResult.Nodes.AddRange(nextListNodeResult.Nodes);
                            isTargetReach = listNodeResult.Nodes.Count >= target;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                foreach (Rest_Node node in listNodeResult.Nodes)
                {
                    AddNodeToEmailInfos(ref emailInfos, node);
                }
            }
            RetrieveEmailInfos_AdditionalInfo(includeConversationID, ref emailInfos);

            result.EmailInfos = emailInfos;
            result.SortedBy = paramSortedBy;//for feedback
            result.SortDirection = sortDirection;//for feedback
            result.MaxCount = maxCount;
            result.TotalCount = listNodeResult.TotalCount;
            result.IsMaxCountReached = result.TotalCount > maxCount;

            // Sort the emails
            result.EmailInfos = sortEmailInfos(paramSortedBy, sortDirection, result.EmailInfos);

            // Comment for now, the grouping is done by frontend side.
            // List<EmailInfo> groupedEmailInfos = result.EmailInfos.GroupBy(x => x.ConversationId).SelectMany(gr => gr).ToList();
            // result.EmailInfos = groupedEmailInfos;

            return result;
        }

        public static ListEmailsResultVer2 ListEmailsVer2(long folderId, string userNameToImpersonate, int pageNumber, int pageSize, string paramSortedBy, string sortDirection, bool showAsConversation)
        {
            ListEmailsResultVer2 result = new ListEmailsResultVer2();
            List<EmailInfoVer2> emailInfos = new List<EmailInfoVer2>();
            string sortColumn = GetSortByColumnVer2(paramSortedBy);

            int totalPage;
            int totalCount;

            long userID = AGOServices.GetCurrentUserID(userNameToImpersonate);
            emailInfos = AGODataAccess.GetEmails(showAsConversation, userID, folderId, pageNumber, pageSize, sortColumn, sortDirection, out totalPage, out totalCount);

            if (showAsConversation)
            {
                foreach (var email in emailInfos)
                {
                    email.Conversations = AGODataAccess.GetConversations(email.NodeID, userID);
                }

                result.EmailInfos = emailInfos;
                result.SortedBy = paramSortedBy;
                result.SortDirection = sortDirection;
                result.TotalCount = totalCount;
                result.PageNumber = pageNumber;
                result.PageSize = pageSize;
                result.TotalPage = totalPage;
            }
            else
            {
                result.EmailInfos = emailInfos;
                result.SortedBy = paramSortedBy;
                result.SortDirection = sortDirection;
                result.TotalCount = totalCount;
                result.PageNumber = pageNumber;
                result.PageSize = pageSize;
                result.TotalPage = totalPage;
            }

            return result;
        }

        private static List<EmailInfo> sortEmailInfos(string paramSortedBy, string sortDirection, List<EmailInfo> emailInfos)
        {
            if (paramSortedBy == "OTEmailFrom")
            {
                if (sortDirection == "asc")
                {
                    // emailInfos.Sort((EmailInfo x, EmailInfo y) => x.EmailFrom.CompareTo(y.EmailFrom));
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.EmailFrom).ToList();
                    emailInfos = sortedEmailInfos;
                }
                else if (sortDirection == "desc")
                {
                    // emailInfos.Sort((EmailInfo x, EmailInfo y) => x.EmailFrom.CompareTo(y.EmailFrom));
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.EmailFrom).ToList();
                    emailInfos = sortedEmailInfos;
                    emailInfos.Reverse();
                }
            }
            else if (paramSortedBy == "OTEmailSubject")
            {
                if (sortDirection == "asc")
                {
                    // emailInfos.Sort((EmailInfo x, EmailInfo y) => x.EmailSubject.CompareTo(y.EmailSubject));
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.EmailSubject).ToList();
                    emailInfos = sortedEmailInfos;
                }
                else if (sortDirection == "desc")
                {
                    // emailInfos.Sort((EmailInfo x, EmailInfo y) => x.EmailSubject.CompareTo(y.EmailSubject));
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.EmailSubject).ToList();
                    emailInfos = sortedEmailInfos;
                    emailInfos.Reverse();
                }
            }
            else if (paramSortedBy == "OTEmailReceivedDate")
            {
                if (sortDirection == "asc")
                {
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.ReceivedDate).ToList();
                    emailInfos = sortedEmailInfos;
                }
                else if (sortDirection == "desc")
                {
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.ReceivedDate).ToList();
                    emailInfos = sortedEmailInfos;
                    emailInfos.Reverse();
                }
            }
            else if (paramSortedBy == "OTEmailTo")
            {
                if (sortDirection == "asc")
                {
                    // emailInfos.Sort((EmailInfo x, EmailInfo y) => x.EmailTo.CompareTo(y.EmailTo));
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.EmailTo).ToList();
                    emailInfos = sortedEmailInfos;
                }
                else if (sortDirection == "desc")
                {
                    // emailInfos.Sort((EmailInfo x, EmailInfo y) => x.EmailTo.CompareTo(y.EmailTo));
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.EmailTo).ToList();
                    emailInfos = sortedEmailInfos;
                    emailInfos.Reverse();
                }
            }
            else if (paramSortedBy == "OTEmailSentDate")
            {
                if (sortDirection == "asc")
                {
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.SentDate).ToList();
                    emailInfos = sortedEmailInfos;
                }
                else if (sortDirection == "desc")
                {
                    List<EmailInfo> sortedEmailInfos = emailInfos.OrderBy(x => x.SentDate).ToList();
                    emailInfos = sortedEmailInfos;
                    emailInfos.Reverse();
                }
            }
            return emailInfos;
        }

        private static string GetSortByColumn(string paramSortedBy)
        {
            string sortByColumn = "emailreceiveddate";
            // In Facets create as a sortable column, don't need to show.
            if (paramSortedBy != null)
            {
                if (paramSortedBy.Equals("OTEmailCC ", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "emailcc";
                }
                else if (paramSortedBy.Equals("OTEmailFrom", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "emailfrom";
                }
                else if (paramSortedBy.Equals("OTEmailHasAttachments", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "emailhasattach";
                }
                else if (paramSortedBy.Equals("OTEmailReceivedDate", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "emailreceiveddate";
                }
                else if (paramSortedBy.Equals("OTEmailSentDate", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "emailsentdate";
                }
                else if (paramSortedBy.Equals("OTEmailSubject", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "emailsubject";
                }
                else if (paramSortedBy.Equals("OTEmailTo", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "emailto";
                }
            }

            return sortByColumn;
        }

        private static string GetSortByColumnVer2(string paramSortedBy)
        {
            string sortByColumn = "EmailReceivedDate";
            if (paramSortedBy != null)
            {
                if (paramSortedBy.Equals("OTEmailFrom", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "EmailFrom";
                }
                else if (paramSortedBy.Equals("OTEmailReceivedDate", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "EmailReceivedDate";
                }
                else if (paramSortedBy.Equals("OTEmailSentDate", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "EmailSentDate";
                }
                else if (paramSortedBy.Equals("OTEmailSubject", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "EmailSubject";
                }
                else if (paramSortedBy.Equals("OTEmailTo", StringComparison.OrdinalIgnoreCase))
                {
                    sortByColumn = "EmailTo";
                }
            }

            return sortByColumn;
        }

        private static void AddNodeToEmailInfos(ref List<EmailInfo> emailInfos, Node node)
        {
            //node validity is now taken care by listnodebypage, where we can specify the type to include
            //if (IsNodeAValidEmail(node))
            //{
            long fileSize = 0;
            if (node.VersionInfo != null)
            {
                fileSize = (long)node.VersionInfo.FileDataSize;
            }
            emailInfos.Add(new EmailInfo()
            {
                Name = node.Name,
                NodeID = node.ID,
                ParentNodeID = node.ParentID,
                FileSize = fileSize
            });
            //}
        }

        private static void AddNodeToEmailInfos(ref List<EmailInfo> emailInfos, Rest_Node node)
        {
            emailInfos.Add(new EmailInfo()
            {
                Name = node.Name,
                PermSeeContents = node.PermSeeContents,
                NodeID = node.ID,
                ParentNodeID = node.ParentID,
                FileSize = node.Size
            });
            foreach (var email in emailInfos)
            {
                logger.Info($"emailInfos PermSeeContents {email.PermSeeContents}");

            }
        }

        private static void AddNodeToEmailInfosVer2(ref List<EmailInfoVer2> emailInfos, Rest_NodeVer2 node)
        {
            emailInfos.Add(new EmailInfoVer2()
            {
                Name = node.Name,
                NodeID = node.Id,
                ParentNodeID = node.ParentId,
                FileSize = node.Size,
                ConversationId = node.ConversationId,
                EmailSubject = node.EmailSubject,
                EmailTo = node.EmailTo,
                EmailCC = node.EmailCc,
                EmailFrom = node.EmailFrom,
                SentDate = node.SentDate,
                ReceivedDate = node.ReceivedDate,
                HasAttachments = node.HasAttachments
            });
        }

        private static void RetrieveEmailInfos_AdditionalInfo(bool includeConversationID, ref List<EmailInfo> emailInfos)
        {
            Stopwatch stopwatch2 = new Stopwatch();
            stopwatch2.Start();

            for (int i = 0; i < emailInfos.Count; i++)
            {
                try
                {
                    EmailInfo emailInfo = emailInfos[i];
                    AGODataAccess.GetEmailInfoFromCSDB(ref emailInfo);

                    if (includeConversationID)
                    {
                        emailInfo.ConversationId = AGODataAccess.GetEmailConversationIDFromCSDB(emailInfo.NodeID);
                    }
                    emailInfos[i] = emailInfo;
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ex.StackTrace);
                }
            }

            stopwatch2.Stop();

            TimeSpan elapsedTime = stopwatch2.Elapsed;
            logger.Info($"Time taken to retrieve additional email info: {elapsedTime.TotalMilliseconds} ms");
        }

        private static void RetrieveEmailAdditionalInfo(bool includeConversationID, ref List<EmailInfo> emailInfos)
        {
            for (int i = 0; i < emailInfos.Count; i++)
            {
                try
                {
                    EmailInfo emailInfo = emailInfos[i];
                    AGODataAccess.GetEmailInfoFromCSDB(ref emailInfo);

                    if (includeConversationID)
                    {
                        emailInfo.ConversationId = AGODataAccess.GetEmailConversationIdStoredProcedure(emailInfo.NodeID);
                    }
                    emailInfos[i] = emailInfo;
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ex.StackTrace);
                }
            }
        }

        private static bool IsNodeAValidEmail(Node node)
        {
            return node.IsContainer == false
                                    && CSAccess.IsNodeHidden(node) == false
                                    && node.Type.Equals("Email", StringComparison.OrdinalIgnoreCase);
        }



        public static ListEmailsResult FindConversationEmails(string userNameToImpersonate, long specifiedEmailID, string sortedBy, string sortDirection)
        {
            ListEmailsResult result = new ListEmailsResult();
            List<EmailInfo> emailInfos = new List<EmailInfo>();
            int maxCount = Properties.Settings.Default.FlatEmailListing_MaxCount;

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            try
            {
                if (specifiedEmailID > 0)
                {
                    long userID = AGOServices.GetCurrentUserID(userNameToImpersonate);
                    emailInfos = AGODataAccess.GetConversations(specifiedEmailID, userID);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ex.StackTrace);
            }

            result.EmailInfos = emailInfos;
            result.MaxCount = maxCount;
            result.IsMaxCountReached = emailInfos.Count > maxCount;
            result.SortedBy = sortedBy;
            result.SortDirection = sortDirection;

            stopwatch.Stop();

            TimeSpan elapsedTime = stopwatch.Elapsed;
            logger.Info($"Time taken to get email conversations: {elapsedTime.TotalMilliseconds} ms");

            return result;
        }
        public static void SortEmailInfos(ref List<EmailInfo> emailInfos, string paramSortedBy, string sortDirection)
        {
            if (paramSortedBy.Equals("OTEmailCC ", StringComparison.OrdinalIgnoreCase))
            {
                if (sortDirection.Equals("asc", StringComparison.OrdinalIgnoreCase))
                {
                    emailInfos.Sort((x, y) => x.EmailCC.CompareTo(y.EmailCC));
                }
                else
                {
                    //default sort is desc
                    emailInfos.Sort((x, y) => y.EmailCC.CompareTo(x.EmailCC));
                }
            }
            else if (paramSortedBy.Equals("OTEmailFrom", StringComparison.OrdinalIgnoreCase))
            {
                if (sortDirection.Equals("asc", StringComparison.OrdinalIgnoreCase))
                {
                    emailInfos.Sort((x, y) => x.EmailFrom.CompareTo(y.EmailFrom));
                }
                else
                {
                    emailInfos.Sort((x, y) => y.EmailFrom.CompareTo(x.EmailFrom));
                }
            }
            else if (paramSortedBy.Equals("OTEmailHasAttachments", StringComparison.OrdinalIgnoreCase))
            {
                if (sortDirection.Equals("asc", StringComparison.OrdinalIgnoreCase))
                {
                    emailInfos.Sort((x, y) => x.HasAttachments.CompareTo(y.HasAttachments));
                }
                else
                {
                    emailInfos.Sort((x, y) => y.HasAttachments.CompareTo(x.HasAttachments));
                }
            }
            else if (paramSortedBy.Equals("OTEmailSentDate", StringComparison.OrdinalIgnoreCase))
            {
                if (sortDirection.Equals("asc", StringComparison.OrdinalIgnoreCase))
                {
                    emailInfos.Sort((x, y) => x.SentDate.GetValueOrDefault().CompareTo(y.SentDate.GetValueOrDefault()));
                }
                else
                {
                    emailInfos.Sort((x, y) => y.SentDate.GetValueOrDefault().CompareTo(x.SentDate.GetValueOrDefault()));
                }
            }
            else if (paramSortedBy.Equals("OTEmailSubject", StringComparison.OrdinalIgnoreCase))
            {
                if (sortDirection.Equals("asc", StringComparison.OrdinalIgnoreCase))
                {
                    emailInfos.Sort((x, y) => x.EmailSubject.CompareTo(y.EmailSubject));
                }
                else
                {
                    emailInfos.Sort((x, y) => y.EmailSubject.CompareTo(x.EmailSubject));
                }
            }
            else if (paramSortedBy.Equals("OTEmailTo", StringComparison.OrdinalIgnoreCase))
            {
                if (sortDirection.Equals("asc", StringComparison.OrdinalIgnoreCase))
                {
                    emailInfos.Sort((x, y) => x.EmailTo.CompareTo(y.EmailTo));
                }
                else
                {
                    emailInfos.Sort((x, y) => y.EmailTo.CompareTo(x.EmailTo));
                }
            }
            else
            {
                //OTEmailReceivedDate is the default sort
                if (sortDirection.Equals("asc", StringComparison.OrdinalIgnoreCase))
                {
                    emailInfos.Sort((x, y) => x.ReceivedDate.GetValueOrDefault().CompareTo(y.ReceivedDate.GetValueOrDefault()));
                }
                else
                {
                    emailInfos.Sort((x, y) => y.ReceivedDate.GetValueOrDefault().CompareTo(x.ReceivedDate.GetValueOrDefault()));
                }
            }
        }
    }
}