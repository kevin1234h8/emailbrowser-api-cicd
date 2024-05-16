using AGOServer.Components;
using AGOServer.Components.Models.OpenText;
using CSAccess_Local;
using OpenText.Livelink.Service.SearchServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class EmailSearchService
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static string GetSortByColumn(string paramSortedBy)
        {
            //In Search Manager, mark as Sortable

            string sortByColumn = SearchResultField.ReceivedDate.Value;//Default sort
            HashSet<string> sortableColumns = new HashSet<string>();
            sortableColumns.Add(SearchResultField.From.Value);
            sortableColumns.Add(SearchResultField.SentDate.Value);
            sortableColumns.Add(SearchResultField.To.Value);
            sortableColumns.Add(SearchResultField.ReceivedDate.Value);
            sortableColumns.Add(SearchResultField.CC.Value);
            sortableColumns.Add(SearchResultField.HasAttachments.Value);
            sortableColumns.Add(SearchResultField.Subject.Value);

            if (paramSortedBy != null)
            {
                if (sortableColumns.Contains(paramSortedBy))
                {
                    sortByColumn = paramSortedBy;
                }
            }
            return sortByColumn;
        }
        public static EmailSearchResult SearchForEmails(string userNameToImpersonate, string keywords, long folderNodeID,
            long? folderNodeID2, long? folderNodeID3, long? folderNodeID4,
            int firstResultToRetrieve, int numResultsToRetrieve, string paramSortedBy, string sortDirection,
            string OTEmailReceivedDate_From, string OTEmailReceivedDate_To, SearchFilterByAttachment searchFilterByAttachment,
            string OTEmailSubject, string OTEmailFrom, string OTEmailTo, bool includeConversationID, string attachmentFileName, string cacheId)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            EmailSearchResult emailSearchResult = null;

            string sortByColumn = GetSortByColumn(paramSortedBy);
            int pageNumber = (int)(Math.Ceiling((float)(firstResultToRetrieve-1) /numResultsToRetrieve))+1;

            Dictionary<string, string> searchOptions = new Dictionary<string, string>() {
                { "page", pageNumber.ToString() },
                { "limit", numResultsToRetrieve.ToString() },
                { "keywords", (!string.IsNullOrEmpty(keywords)) ? keywords.Replace(",", "\\,").Replace("!", "\\!").Replace(":", "\\:").Replace("'", "\\'").Replace("[", "\\[").Replace("]", "\\]").Replace("(", "\\(").Replace(")", "\\)") : "" }, //
                { "SortColumn", sortByColumn },
                { "SortDirection", sortDirection },
                { "OTEmailSubject", (!string.IsNullOrEmpty(OTEmailSubject)) ? OTEmailSubject.Replace(",", "\\,").Replace("!", "\\!").Replace(":", "\\:").Replace("'", "\\'").Replace("[", "\\[").Replace("]", "\\]").Replace("(", "\\(").Replace(")", "\\)") : "" },
                { "OTEmailSentDate_Start", OTEmailReceivedDate_From },
                { "OTEmailSentDate_End", OTEmailReceivedDate_To },
                { "OTEmailReceivedDate_Start", "" },
                { "OTEmailReceivedDate_End", "" },
                { "OTEmailFrom", OTEmailFrom },
                { "OTEmailTo", OTEmailTo },
                { "SearchByAttachment", ((int)searchFilterByAttachment).ToString() },
                { "FolderNodeID", folderNodeID.ToString() },
                { "FolderNodeID2", (folderNodeID2 != null) ? folderNodeID2.ToString() : "" },
                { "FolderNodeID3", (folderNodeID3 != null) ? folderNodeID3.ToString() : "" },
                { "FolderNodeID4", (folderNodeID4 != null) ? folderNodeID4.ToString() : "" },
                { "AttachmentName", (!string.IsNullOrEmpty(attachmentFileName)) ? attachmentFileName.Replace(",", "\\,").Replace("!", "\\!").Replace(":", "\\:").Replace("'", "\\'").Replace("[", "\\[").Replace("]", "\\]").Replace("(", "\\(").Replace(")", "\\)") : "" },
                { "CacheID", cacheId }
            };
            logger.Info($"Search Keyword: {keywords}");
            OpenTextV2SearchResponse searchV2Response = null;

            Stopwatch stopwatch2 = new Stopwatch();
            stopwatch2.Start();

            if (string.IsNullOrEmpty(attachmentFileName))
            {
                if (Properties.Settings.Default.SearchMode.ToLower() == "api")
                {
                    searchV2Response = CSAccess.Rest_Search_Fuzzy(searchOptions, userNameToImpersonate);
                }
                else
                {
                    searchV2Response = CSAccess.Rest_Search_Form(searchOptions, userNameToImpersonate);
                }
            }
            else
            {
                searchV2Response = CSAccess.Rest_Search_Fuzzy(searchOptions, userNameToImpersonate);
            }

            stopwatch2.Stop();
            TimeSpan elapsedTime2 = stopwatch2.Elapsed;
            logger.Info($"Time taken to get search result from API: {elapsedTime2.TotalMilliseconds} ms");

            //In Search Manager, mark as displayable
            string[] resultColumns = new string[] {
                SearchResultField.OTLocation.Value,
                SearchResultField.FileSize.Value,
                SearchResultField.OTName.Value,
                SearchResultField.Summary.Value
            };

            //SingleSearchResponse searchResponse = CSAccess.Search_Fuzzy(userNameToImpersonate, keywords, folderNodeID,
            //    folderNodeID2, folderNodeID3, folderNodeID4,
            //    firstResultToRetrieve, numResultsToRetrieve, resultColumns, sortByColumn, sortDirection,
            // OTEmailReceivedDate_From, OTEmailReceivedDate_To, searchFilterByAttachment,
            // OTEmailSubject, OTEmailFrom, OTEmailTo);
            logger.Trace("Sortbycolumn: " + sortByColumn);

            try
            {
                if (searchV2Response != null)
                {
                    //SResultPage page = searchResponse.Results;

                    //logger.Trace("searchFilterByAttachment: " + searchFilterByAttachment);
                    //emailSearchResult = Parse_CS_SearchResponse(searchResponse);
                    emailSearchResult = ParseCSRestSearchResponse(searchV2Response);

                    if (emailSearchResult != null && emailSearchResult.EmailInfos != null)
                    {
                        logger.Trace("Email Infos from search result count: " + emailSearchResult.EmailInfos.Count);
                        // iterate emails list in reverse, so the email can be deleted easily from list when necessary
                        int numberOfEmails = emailSearchResult.EmailInfos.Count;
                        for (int i = numberOfEmails - 1; i >= 0; i--)
                        {
                            EmailInfo emailInfo = emailSearchResult.EmailInfos[i];
                            AGODataAccess.GetEmailInfoFromCSDB(ref emailInfo);

                            if (includeConversationID)
                            {
                                // Add ConversationId value to emailInfo object
                                if (!AGODataAccess.IsFromMigratedFolder(emailInfo.NodeID.ToString()))
                                    emailInfo.ConversationId = AGODataAccess.GetEmailConversationIDFromCSDB(emailInfo.NodeID);
                                emailInfo.ClientLocalEmailID = AGODataAccess.GetLocalEmailIDfromEmailFilling(emailInfo.NodeID);
                            }
                            // Add ClientLocalEmailID value to emailInfo
                            //emailInfo.ClientLocalEmailID = AGODataAccess.GetLocalEmailIDfromEmailFilling(emailInfo.NodeID);

                            // if the attachmentFileName parameter is neither empty nor null then the emails must have attachment
                            // and searchFilterByAttachment parameter must have "YES" value

                            /*if (attachmentFileName != "" && attachmentFileName != null && emailInfo.HasAttachments == 1) // advance search
                                                                                                                         // || searchFilterByAttachment.ToString() == "YES" || searchFilterByAttachment.ToString() == "NO") 
                            {
                                logger.Trace("This is advance search");
                                emailInfo.Attachments = AttachmentServices.ListAttachments(userNameToImpersonate, emailInfo.NodeID);
                                // even though emailInfo.HasAttachments == 1, an email might have no attachment in it
                                // so we must filter the emails again
                                if (emailInfo.Attachments.Count > 0)
                                {
                                    logger.Trace("EmailID = " + emailInfo.NodeID + " | Attachments = " + emailInfo.Attachments.Count
                                        + " | attachmentFileName = " + attachmentFileName + " | FirstAttachment = "
                                        + emailInfo.Attachments[0].FileName);
                                    bool match = false;
                                    for (int n = 0; n < emailInfo.Attachments.Count; n++)
                                    {
                                        // check whether the email's attachment file name matches the attachmentFileName parameter
                                        bool comparisonResult = emailInfo.Attachments[n].FileName.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0;
                                        if (comparisonResult)
                                        {
                                            match = true;
                                            logger.Trace(emailInfo.NodeID + " matched");
                                            // it seems that the emailInfo is passed by reference to emailSearchResult.EmailInfos[i], so this is not required
                                            // emailSearchResult.EmailInfos[i] = emailInfo;
                                            break;
                                        }
                                    }
                                    // if the email doesn't have any attachment whose file name matches the attachmentFileName parameter
                                    // then remove the email from the list
                                    if (!match)
                                    {
                                        logger.Trace("Not matched! Remove: " + emailSearchResult.EmailInfos[i].NodeID);
                                        emailSearchResult.EmailInfos.RemoveAt(i);
                                    }
                                }
                                else // if emailInfo.Attachments.Count == 0, remove the email from search result
                                {
                                    logger.Trace("Does not contain attachment! Remove: " + emailSearchResult.EmailInfos[i].NodeID);
                                    emailSearchResult.EmailInfos.RemoveAt(i);
                                }
                            }
                            else // normal search
                            {
                                logger.Trace("This is normal search");
                            }*/

                            // emailSearchResult.EmailInfos[i] = emailInfo; // not required because emailInfo is passed by reference to emailSearchResult.EmailInfos[i]
                        }
                    }
                    else
                    {
                        logger.Trace("parsed Email Infos from search result is null");
                    }
                }
                else
                {
                    logger.Trace("the search response is null");
                }

                // update IncludeCount (number of emails) because some emails might be removed
                // because they don't have attachment but their HasAttachments property value is 1
                // or because each of them doesn't have any attachment whose file name matches the attachmentFileName parameter
                emailSearchResult.IncludeCount = emailSearchResult.EmailInfos.Count;




                // Sort search result by received date
                // The Search_Fuzzy sorting-by-received-date doesn't work so we make our own
                if (sortByColumn == "OTEmailReceivedDate")
                {
                    if (sortDirection == "asc")
                    {
                        List<EmailInfo> sortedEmailInfos = emailSearchResult.EmailInfos.OrderBy(x => x.ReceivedDate).ToList();
                        emailSearchResult.EmailInfos = sortedEmailInfos;
                    }
                    else if (sortDirection == "desc")
                    {
                        List<EmailInfo> sortedEmailInfos = emailSearchResult.EmailInfos.OrderBy(x => x.ReceivedDate).ToList();
                        emailSearchResult.EmailInfos = sortedEmailInfos;
                        emailSearchResult.EmailInfos.Reverse();
                    }
                }

                if (sortByColumn == "OTEmailSentDate")
                {
                    if (sortDirection == "asc")
                    {
                        List<EmailInfo> sortedEmailInfos = emailSearchResult.EmailInfos.OrderBy(x => x.SentDate).ToList();
                        emailSearchResult.EmailInfos = sortedEmailInfos;
                    }
                    else if (sortDirection == "desc")
                    {
                        List<EmailInfo> sortedEmailInfos = emailSearchResult.EmailInfos.OrderBy(x => x.SentDate).ToList();
                        emailSearchResult.EmailInfos = sortedEmailInfos;
                        emailSearchResult.EmailInfos.Reverse();
                    }
                }

                // Group emails by ConversationId
                //List<EmailInfo> groupedEmailInfos = emailSearchResult.EmailInfos.GroupBy(x => x.ConversationId).SelectMany(gr => gr).ToList();
                //emailSearchResult.EmailInfos = groupedEmailInfos;

                /**
                // Remove identical emails that are located in different folder
                for (int i = 0; i < emailSearchResult.EmailInfos.Count; i++)
                {
                    // Search emails that have the same ConversationIds but have different ParentNodeIDs (located in different folders),
                    // then remove them from emails list
                    EmailInfo currentEmail = emailSearchResult.EmailInfos[i];
                    for (int j = 0; j < emailSearchResult.EmailInfos.Count; j++)
                    {
                        // if not in the same group
                        if (currentEmail.ConversationId != emailSearchResult.EmailInfos[j].ConversationId)
                        {
                            break; // exit the loop and move to the next email
                        }

                        // if in the same group but different folder
                        if (currentEmail.ConversationId == emailSearchResult.EmailInfos[j].ConversationId
                            && currentEmail.ParentNodeID != emailSearchResult.EmailInfos[j].ParentNodeID)
                        {
                            groupedEmailInfos.RemoveAt(j);
                        }
                    }
                }

                // update EmailInfos
                emailSearchResult.EmailInfos = groupedEmailInfos;

                // update ActualCount and IncludeCount
                emailSearchResult.ActualCount = emailSearchResult.EmailInfos.Count;
                emailSearchResult.IncludeCount = emailSearchResult.EmailInfos.Count;
                **/
            }
            catch (Exception ex)
            {
                logger.Error("error in Search :" + ex.Message + ex.StackTrace);
            }

            stopwatch.Stop();

            TimeSpan elapsedTime = stopwatch.Elapsed;
            logger.Info($"Time taken to get search result: {elapsedTime.TotalMilliseconds} ms");

            return emailSearchResult;
        }

        public static EmailSearchResultVer2 SearchForEmailsVer2(string userNameToImpersonate, string keywords, long folderNodeID,
            long? folderNodeID2, long? folderNodeID3, long? folderNodeID4,
            int firstResultToRetrieve, int numResultsToRetrieve, string paramSortedBy, string sortDirection,
            string OTEmailReceivedDate_From, string OTEmailReceivedDate_To, SearchFilterByAttachment searchFilterByAttachment,
            string OTEmailSubject, string OTEmailFrom, string OTEmailTo, bool includeConversationID, string attachmentFileName, int pageNumber, int pageSize)
        {
            EmailSearchResultVer2 emailSearchResult = new EmailSearchResultVer2();
            int maxCount = Properties.Settings.Default.FlatEmailListing_MaxCount;
            string sortByColumn = GetSortByColumn(paramSortedBy);
            int originalCount = 0;

            // In search manager mark as displayable.
            string[] resultColumns = new string[] {
                SearchResultField.OTLocation.Value,
                SearchResultField.FileSize.Value,
                SearchResultField.OTName.Value,
                SearchResultField.Summary.Value
            };

            SingleSearchResponse searchResponse = CSAccess.Search_Fuzzy(userNameToImpersonate, keywords, folderNodeID,
                folderNodeID2, folderNodeID3, folderNodeID4,
                firstResultToRetrieve, maxCount, resultColumns, sortByColumn, sortDirection,
             OTEmailReceivedDate_From, OTEmailReceivedDate_To, searchFilterByAttachment,
             OTEmailSubject, OTEmailFrom, OTEmailTo);

            logger.Trace("Sortbycolumn: " + sortByColumn);

            if (searchResponse != null)
            {
                SResultPage page = searchResponse.Results;

                logger.Trace("searchFilterByAttachment: " + searchFilterByAttachment);
                emailSearchResult = Parse_CS_SearchResponseVer2(searchResponse);
                originalCount = emailSearchResult.EmailInfos.Count;
                if (emailSearchResult != null && emailSearchResult.EmailInfos != null)
                {
                    logger.Trace("Email Infos from search result count: " + emailSearchResult.EmailInfos.Count);
                    // iterate emails list in reverse, so the email can be deleted easily from list when necessary
                    int numberOfEmails = emailSearchResult.EmailInfos.Count;
                    for (int i = numberOfEmails - 1; i >= 0; i--)
                    {
                        EmailSearchInfoVer2 emailInfo = emailSearchResult.EmailInfos[i];
                        AGODataAccess.GetEmailInfoFromCSDBVer2(ref emailInfo);

                        if (includeConversationID)
                        {
                            // Add ConversationId value to emailInfo object
                            emailInfo.ConversationId = AGODataAccess.GetEmailConversationIDFromCSDB(emailInfo.NodeID);
                            emailInfo.ClientLocalEmailID = AGODataAccess.GetLocalEmailIDfromEmailFilling(emailInfo.NodeID);
                        }
                        // Add ClientLocalEmailID value to emailInfo
                        //emailInfo.ClientLocalEmailID = AGODataAccess.GetLocalEmailIDfromEmailFilling(emailInfo.NodeID);

                        // if the attachmentFileName parameter is neither empty nor null then the emails must have attachment
                        // and searchFilterByAttachment parameter must have "YES" value
                        if (attachmentFileName != "" && attachmentFileName != null && emailInfo.HasAttachments == 1) // advance search
                                                                                                                     // || searchFilterByAttachment.ToString() == "YES" || searchFilterByAttachment.ToString() == "NO") 
                        {
                            logger.Trace("This is advance search");
                            emailInfo.Attachments = AttachmentServices.ListAttachments(userNameToImpersonate, emailInfo.NodeID);
                            // even though emailInfo.HasAttachments == 1, an email might have no attachment in it
                            // so we must filter the emails again
                            if (emailInfo.Attachments.Count > 0)
                            {
                                logger.Trace("EmailID = " + emailInfo.NodeID + " | Attachments = " + emailInfo.Attachments.Count
                                    + " | attachmentFileName = " + attachmentFileName + " | FirstAttachment = "
                                    + emailInfo.Attachments[0].FileName);
                                bool match = false;
                                for (int n = 0; n < emailInfo.Attachments.Count; n++)
                                {
                                    // check whether the email's attachment file name matches the attachmentFileName parameter
                                    bool comparisonResult = emailInfo.Attachments[n].FileName.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0;
                                    if (comparisonResult)
                                    {
                                        match = true;
                                        logger.Trace(emailInfo.NodeID + " matched");
                                        // it seems that the emailInfo is passed by reference to emailSearchResult.EmailInfos[i], so this is not required
                                        // emailSearchResult.EmailInfos[i] = emailInfo;
                                        break;
                                    }
                                }
                                // if the email doesn't have any attachment whose file name matches the attachmentFileName parameter
                                // then remove the email from the list
                                if (!match)
                                {
                                    logger.Trace("Not matched! Remove: " + emailSearchResult.EmailInfos[i].NodeID);
                                    emailSearchResult.EmailInfos.RemoveAt(i);
                                }
                            }
                            else // if emailInfo.Attachments.Count == 0, remove the email from search result
                            {
                                logger.Trace("Does not contain attachment! Remove: " + emailSearchResult.EmailInfos[i].NodeID);
                                emailSearchResult.EmailInfos.RemoveAt(i);
                            }
                        }
                        else // normal search
                        {
                            logger.Trace("This is normal search");
                        }
                        // emailSearchResult.EmailInfos[i] = emailInfo; // not required because emailInfo is passed by reference to emailSearchResult.EmailInfos[i]
                    }
                }
                else
                {
                    logger.Trace("parsed Email Infos from search result is null");
                }
            }
            else
            {
                logger.Trace("the search response is null");
            }

            // update IncludeCount (number of emails) because some emails might be removed
            // because they don't have attachment but their HasAttachments property value is 1
            // or because each of them doesn't have any attachment whose file name matches the attachmentFileName parameter
            emailSearchResult.IncludeCount = emailSearchResult.EmailInfos.Count;

            var emailInfos = emailSearchResult.EmailInfos;
            logger.Info("emailInfos: " + emailInfos.Count());
            var groupedEmailInfos = emailInfos.GroupBy(x => x.ConversationId)
                    .Select(group => new { ConversationId = group.Key, Conversations = group.OrderByDescending(o => o.SentDate) });
            logger.Info("groupedEmailInfos: " + groupedEmailInfos.Count());

            #region Sorting by emailFrom, emailTo, receivedDate and sendDate.
            if (paramSortedBy == "OTEmailFrom")
            {
                if (sortDirection == "asc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderBy(x => x.Conversations.First().EmailFrom);
                }
                else if (sortDirection == "desc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderByDescending(x => x.Conversations.First().EmailFrom);
                }
            }
            else if (paramSortedBy == "OTEmailSubject")
            {
                if (sortDirection == "asc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderBy(x => x.Conversations.First().EmailSubject);
                }
                else if (sortDirection == "desc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderByDescending(x => x.Conversations.First().EmailSubject);
                }
            }
            else if (paramSortedBy == "OTEmailReceivedDate")
            {
                if (sortDirection == "asc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderBy(x => x.Conversations.First().ReceivedDate);
                }
                else if (sortDirection == "desc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderByDescending(x => x.Conversations.First().ReceivedDate);
                }
            }
            else if (paramSortedBy == "OTEmailTo")
            {
                if (sortDirection == "asc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderBy(x => x.Conversations.First().EmailTo);
                }
                else if (sortDirection == "desc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderByDescending(x => x.Conversations.First().EmailTo);
                }
            }
            else if (paramSortedBy == "OTEmailSentDate")
            {
                if (sortDirection == "asc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderBy(x => x.Conversations.First().SentDate);
                }
                else if (sortDirection == "desc")
                {
                    groupedEmailInfos = groupedEmailInfos.OrderByDescending(x => x.Conversations.First().SentDate);
                }
            }
            #endregion

            int totalCount = groupedEmailInfos.Count();
            int totalPage = Convert.ToInt32(Math.Ceiling(totalCount / Convert.ToDouble(pageSize)));
            logger.Info("pageSize: " + pageSize);
            logger.Info("pageNumber: " + pageNumber);

            // Pagination using LINQ.
            // Skip function will ignore the first n items. 
            // Take function will limit how many items are taken. 
            groupedEmailInfos = groupedEmailInfos.Skip(pageSize * (pageNumber - 1)).Take(pageSize);
            logger.Info("groupedEmailInfos after pagination: " + groupedEmailInfos.Count());

            List<EmailSearchInfoVer2> newEmailInfos = new List<EmailSearchInfoVer2>();
            foreach (var grp in groupedEmailInfos)
            {
                EmailSearchInfoVer2 newEmailInfo = new EmailSearchInfoVer2();
                newEmailInfo.Name = grp.Conversations.First().Name;
                newEmailInfo.NodeID = grp.Conversations.First().NodeID;
                newEmailInfo.ParentNodeID = grp.Conversations.First().ParentNodeID;
                newEmailInfo.FileSize = grp.Conversations.First().FileSize;
                newEmailInfo.ConversationId = grp.Conversations.First().ConversationId;
                newEmailInfo.EmailSubject = grp.Conversations.First().EmailSubject;
                newEmailInfo.EmailTo = grp.Conversations.First().EmailTo;
                newEmailInfo.EmailCC = grp.Conversations.First().EmailCC;
                newEmailInfo.EmailFrom = grp.Conversations.First().EmailFrom;
                newEmailInfo.SentDate = grp.Conversations.First().SentDate;
                newEmailInfo.ReceivedDate = grp.Conversations.First().ReceivedDate;
                newEmailInfo.HasAttachments = grp.Conversations.First().HasAttachments;

                List<EmailInfo> convs = new List<EmailInfo>();
                foreach (var conversation in grp.Conversations)
                {
                    EmailInfo conv = new EmailInfo();
                    conv.Name = conversation.Name;
                    conv.NodeID = conversation.NodeID;
                    conv.ParentNodeID = conversation.ParentNodeID;
                    conv.FileSize = conversation.FileSize;
                    conv.ConversationId = conversation.ConversationId;
                    conv.EmailSubject = conversation.EmailSubject;
                    conv.EmailTo = conversation.EmailTo;
                    conv.EmailCC = conversation.EmailCC;
                    conv.EmailFrom = conversation.EmailFrom;
                    conv.SentDate = conversation.SentDate;
                    conv.ReceivedDate = conversation.ReceivedDate;
                    conv.HasAttachments = conversation.HasAttachments;
                    convs.Add(conv);
                }
                newEmailInfo.Conversations = convs;

                newEmailInfos.Add(newEmailInfo);
            }

            emailSearchResult.EmailInfos = newEmailInfos;
            emailSearchResult.SortedBy = paramSortedBy;
            emailSearchResult.SortDirection = sortDirection;
            emailSearchResult.TotalCount = totalCount;
            emailSearchResult.IsMaxCountReached = originalCount > maxCount;
            emailSearchResult.PageNumber = pageNumber;
            emailSearchResult.PageSize = pageSize;
            emailSearchResult.TotalPage = totalPage;

            return emailSearchResult;
        }

        public static EmailSearchResultVer2 Parse_CS_SearchResponseVer2(SingleSearchResponse results)
        {
            EmailSearchResultVer2 emailSearchResult = null;
            SResultPage resultPage = results.Results;
            SGraph[] resultAnalysis = results.ResultAnalysis;

            if (resultPage != null)
            {
                // Proceed if there was at least one search result
                if (resultPage.Item != null && resultPage.Type != null && resultPage.Item.Length > 0)
                {
                    logger.Trace("No of search result: " + resultPage.Item.Length);
                    emailSearchResult = ParseSearchResultVer2(results.Type, resultPage, resultAnalysis);
                }
                else
                {
                    logger.Trace("There are no results");
                }
            }
            else
            {
                logger.Trace("resultPage is null");
            }
            return emailSearchResult;
        }

        public static EmailSearchResult Parse_CS_SearchResponse(SingleSearchResponse results)
        {
            EmailSearchResult emailSearchResult = null;
            SResultPage resultPage = results.Results;
            SGraph[] resultAnalysis = results.ResultAnalysis;

            if (resultPage != null)
            {
                // Proceed if there was at least one search result
                if (resultPage.Item != null && resultPage.Type != null && resultPage.Item.Length > 0)
                {
                    logger.Trace("No of search result: " + resultPage.Item.Length);
                    emailSearchResult = ParseSearchResult(results.Type, resultPage, resultAnalysis);
                }
                else
                {
                    logger.Trace("There are no results");
                }
            }
            else
            {
                logger.Trace("resultPage is null");
            }
            return emailSearchResult;
        }

        private static EmailSearchResultVer2 ParseSearchResultVer2(DataBagType[] availableTypes, SResultPage resultPage, SGraph[] resultAnalysis)
        {
            EmailSearchResultVer2 emailSearchResult = new EmailSearchResultVer2();
            SGraph[] items = resultPage.Item;
            DataBagType[] types = resultPage.Type;//This is in the same order as the values
            // Create a key/value lookup for <Region Name, Friendly Display Name>
            Dictionary<string, Dictionary<string, string>> regiontypes = new Dictionary<string, Dictionary<string, string>>();
            SNode regionnode = resultAnalysis[0].N[0];   // standard Content Server only has one Result Analysis graph containing one node

            for (int i = 0; i < types.Length; i++)
            {
                Dictionary<string, string> regions = new Dictionary<string, string>();
                for (int j = 0; j < regionnode.S.Length; j++)
                {
                    regions.Add(availableTypes[i].Strings[j], regionnode.S[j]);
                }
                regiontypes.Add(types[i].ID, regions);
            }

            emailSearchResult.ListHead = resultPage.ListDescription.ListHead;
            emailSearchResult.IncludeCount = resultPage.ListDescription.IncludeCount;
            emailSearchResult.ActualCount = resultPage.ListDescription.ActualCount;
            logger.Trace("ListHead [{0}] IncludeCount[{1}] ActualCount[{2}]", resultPage.ListDescription.ListHead,
                resultPage.ListDescription.IncludeCount, resultPage.ListDescription.ActualCount);
            emailSearchResult.Remarks = String.Format("Search Results: {0} to {1} of about {2}", resultPage.ListDescription.ListHead,
                resultPage.ListDescription.IncludeCount + resultPage.ListDescription.ListHead, resultPage.ListDescription.ActualCount);

            DataBagType type = types[0];
            Dictionary<string, string> regionnames = regiontypes[type.ID];

            // Iterate through the search hits (SGraphs)
            int itemcount = 0;
            foreach (SGraph item in items)
            {
                itemcount++;
                string[] itemData = item.ID.Split(new char[] { '&', '=' });

                long nodeId = -1;
                if (itemData.Length >= 2)
                {
                    string nodeIdStr = itemData[1];
                    long.TryParse(nodeIdStr, out nodeId);
                }

                if (nodeId > 0)
                {
                    EmailSearchInfoVer2 emailInfo = new EmailSearchInfoVer2();
                    emailInfo.NodeID = nodeId;

                    SNode[] SNodes = item.N;
                    string friendlyrgn = "";
                    string originalrgn = "";

                    /* Note: In standard Content Server 9.7.1, there will only be one node per item.
                           This translates to one search hit per OTURN (DataID + VersionID + VerType).
                           Each OTURN is unique within Content Server 9.7.1.
                           With this in mind, the following for loop is optional.
                           Directly accessing nodes[0] would suffice. */
                    //for (int hitNode = 0; hitNode < nodes.Length; hitNode++)
                    //{
                    //    SNode searchNode = nodes[i];
                    //}
                    SNode searchNode = SNodes[0];
                    EmailSearchResultItem resultItem = RetrieveEmailSearchEntry(type, regionnames, ref friendlyrgn, ref originalrgn, searchNode);
                    emailInfo.Summary = resultItem.Summary;

                    if (long.TryParse(resultItem.FileSize, out long emailFileSize))
                    {
                        emailInfo.FileSize = emailFileSize;
                    }
                    if (string.IsNullOrWhiteSpace(resultItem.FullLocation) == false)
                    {
                        string[] emailFullLocation = resultItem.FullLocation.Split(' ');
                        if (emailFullLocation.Length >= 2)
                        {
                            string emailContainer = emailFullLocation[emailFullLocation.Length - 2];
                            if (long.TryParse(emailContainer, out long emailParentNodeID))
                            {
                                emailInfo.ParentNodeID = emailParentNodeID;
                            }
                        }
                    }
                    emailInfo.Name = resultItem.Name;
                    emailSearchResult.EmailInfos.Add(emailInfo);
                    //we will do post process down stream to fill in the rest of the informations.
                }
            }

            return emailSearchResult;
        }

        private static EmailSearchResult ParseSearchResult(DataBagType[] availableTypes, SResultPage resultPage, SGraph[] resultAnalysis)
        {
            EmailSearchResult emailSearchResult = new EmailSearchResult();
            SGraph[] items = resultPage.Item;
            DataBagType[] types = resultPage.Type;//This is in the same order as the values
            // Create a key/value lookup for <Region Name, Friendly Display Name>
            Dictionary<string, Dictionary<string, string>> regiontypes = new Dictionary<string, Dictionary<string, string>>();
            SNode regionnode = resultAnalysis[0].N[0];   // standard Content Server only has one Result Analysis graph containing one node

            for (int i = 0; i < types.Length; i++)
            {
                Dictionary<string, string> regions = new Dictionary<string, string>();
                for (int j = 0; j < regionnode.S.Length; j++)
                {
                    regions.Add(availableTypes[i].Strings[j], regionnode.S[j]);
                }
                regiontypes.Add(types[i].ID, regions);
            }

            emailSearchResult.ListHead = resultPage.ListDescription.ListHead;
            emailSearchResult.IncludeCount = resultPage.ListDescription.IncludeCount;
            emailSearchResult.ActualCount = resultPage.ListDescription.ActualCount;
            logger.Trace("ListHead [{0}] IncludeCount[{1}] ActualCount[{2}]", resultPage.ListDescription.ListHead,
                resultPage.ListDescription.IncludeCount, resultPage.ListDescription.ActualCount);
            emailSearchResult.Remarks = String.Format("Search Results: {0} to {1} of about {2}", resultPage.ListDescription.ListHead,
                resultPage.ListDescription.IncludeCount + resultPage.ListDescription.ListHead, resultPage.ListDescription.ActualCount);

            DataBagType type = types[0];
            Dictionary<string, string> regionnames = regiontypes[type.ID];

            // Iterate through the search hits (SGraphs)
            int itemcount = 0;
            foreach (SGraph item in items)
            {
                itemcount++;
                string[] itemData = item.ID.Split(new char[] { '&', '=' });

                long nodeId = -1;
                if (itemData.Length >= 2)
                {
                    string nodeIdStr = itemData[1];
                    long.TryParse(nodeIdStr, out nodeId);
                }

                if (nodeId > 0)
                {
                    EmailInfo emailInfo = new EmailInfo();
                    emailInfo.NodeID = nodeId;

                    SNode[] SNodes = item.N;
                    string friendlyrgn = "";
                    string originalrgn = "";

                    /* Note: In standard Content Server 9.7.1, there will only be one node per item.
                           This translates to one search hit per OTURN (DataID + VersionID + VerType).
                           Each OTURN is unique within Content Server 9.7.1.
                           With this in mind, the following for loop is optional.
                           Directly accessing nodes[0] would suffice. */
                    //for (int hitNode = 0; hitNode < nodes.Length; hitNode++)
                    //{
                    //    SNode searchNode = nodes[i];
                    //}
                    SNode searchNode = SNodes[0];
                    EmailSearchResultItem resultItem = RetrieveEmailSearchEntry(type, regionnames, ref friendlyrgn, ref originalrgn, searchNode);
                    emailInfo.Summary = resultItem.Summary;

                    if (long.TryParse(resultItem.FileSize, out long emailFileSize))
                    {
                        emailInfo.FileSize = emailFileSize;
                    }
                    if (string.IsNullOrWhiteSpace(resultItem.FullLocation) == false)
                    {
                        string[] emailFullLocation = resultItem.FullLocation.Split(' ');
                        if (emailFullLocation.Length >= 2)
                        {
                            string emailContainer = emailFullLocation[emailFullLocation.Length - 2];
                            if (long.TryParse(emailContainer, out long emailParentNodeID))
                            {
                                emailInfo.ParentNodeID = emailParentNodeID;
                            }
                        }
                    }
                    emailInfo.Name = resultItem.Name;
                    emailSearchResult.EmailInfos.Add(emailInfo);
                    //we will do post process down stream to fill in the rest of the informations.
                }
            }
            return emailSearchResult;
        }

        private static EmailSearchResultItem RetrieveEmailSearchEntry(DataBagType type, Dictionary<string, string> regionnames, ref string friendlyrgn, ref string originalrgn, SNode searchNode)
        {
            EmailSearchResultItem resultItem = new EmailSearchResultItem();
            try
            {
                // Output the result's string regions with their friendly display name, if available
                if (searchNode.S != null)
                {
                    for (int i = 0; i < searchNode.S.Length; i++)
                    {
                        originalrgn = type.Strings[i];
                        friendlyrgn = regionnames.ContainsKey(originalrgn) ? regionnames[originalrgn] : originalrgn;
                        string value = searchNode.S[i];

                        if (originalrgn.ToUpperInvariant().Equals(SearchResultField.Summary.Value.ToUpperInvariant()))
                        {
                            resultItem.Summary = value;
                        }
                        else if (originalrgn.ToUpperInvariant().Equals(SearchResultField.OTLocation.Value.ToUpperInvariant()))
                        {
                            resultItem.FullLocation = value;
                        }
                        else if (originalrgn.ToUpperInvariant().Equals(SearchResultField.OTName.Value.ToUpperInvariant()))
                        {
                            resultItem.Name = value;
                        }
                        else if (originalrgn.ToUpperInvariant().Equals(SearchResultField.FileSize.Value.ToUpperInvariant()))
                        {
                            resultItem.FileSize = value;
                        }
                    }
                }

                // Output the result's integer fields with their friendly display name, if available
                if (searchNode.I != null)
                {
                    for (int i = 0; i < searchNode.I.Length; i++)
                    {
                        originalrgn = type.Ints[i];
                        friendlyrgn = regionnames.ContainsKey(originalrgn) ? regionnames[originalrgn] : originalrgn;
                        //Console.WriteLine("       " + friendlyrgn + "=" + searchNode.I[i].ToString());
                    }
                }

                // Output the result's date fields with their friendly display name, if available
                if (searchNode.D != null)
                {
                    for (int i = 0; i < searchNode.D.Length; i++)
                    {
                        originalrgn = type.Dates[i];
                        friendlyrgn = regionnames.ContainsKey(originalrgn) ? regionnames[originalrgn] : originalrgn;
                        //Console.WriteLine("       " + friendlyrgn + "=" + searchNode.D[i].ToShortDateString() + " @ " + searchNode.D[i].ToShortTimeString());
                    }
                }

                // Output the result's real fields with their friendly display name, if available
                if (searchNode.R != null)
                {
                    for (int i = 0; i < searchNode.R.Length; i++)
                    {
                        originalrgn = type.Reals[i];
                        friendlyrgn = regionnames.ContainsKey(originalrgn) ? regionnames[originalrgn] : originalrgn;
                        //Console.WriteLine("       " + friendlyrgn + "=" + searchNode.R[i].ToString());
                    }
                }

                /* Use this debug section to see how many strings, integers, dates and reals exist for each search result
                   ------------------------------------------------------------------------------------------------------
                    Console.WriteLine();
                    Console.WriteLine("    Strings=" + (node.S != null ? node.S.Length : 0));
                    Console.WriteLine("   Integers=" + (node.I != null ? node.I.Length : 0));
                    Console.WriteLine("      Dates=" + (node.D != null ? node.D.Length : 0));
                    Console.WriteLine("      Reals=" + (node.R != null ? node.R.Length : 0));
                    Console.WriteLine();
                 */
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ex.StackTrace);
            }
            return resultItem;
        }

        public static EmailSearchResult SearchAttachmentFileName(long userId, string keywords, long folderNodeID,
            long? folderNodeID2, long? folderNodeID3, long? folderNodeID4,
            int firstResultToRetrieve, int numResultsToRetrieve, string paramSortedBy, string sortDirection,
            string OTEmailReceivedDate_From, string OTEmailReceivedDate_To, string OTEmailSubject, string OTEmailFrom, string OTEmailTo, string attachmentFileName)
        {
            EmailSearchResult results = AGODataAccess.SearchAttachmentFileNameFromCSDB(userId, attachmentFileName, firstResultToRetrieve, numResultsToRetrieve, OTEmailFrom, OTEmailSubject, OTEmailTo, OTEmailReceivedDate_From, OTEmailReceivedDate_To, paramSortedBy, sortDirection);

            return results;
        }

        public static EmailSearchResult ParseCSRestSearchResponse(OpenTextV2SearchResponse searchResponse)
        {
            EmailSearchResult result = new EmailSearchResult();
            result.ActualCount = 0;
            result.ListHead = 0;
            result.IncludeCount = 0;
            result.EmailInfos = new List<EmailInfo>();

            try
            {
                if (searchResponse != null)
                {
                    if (searchResponse.collection != null && searchResponse.collection.paging != null && searchResponse.collection.searching != null)
                    {
                        result.ActualCount = searchResponse.collection.paging.total_count;
                        result.ListHead = searchResponse.collection.paging.range_min;
                        result.IncludeCount = searchResponse.collection.paging.limit;
                        result.CacheID = searchResponse.collection.searching.cache_id.ToString();
                    }

                    foreach (var data in searchResponse.results)
                    {
                        EmailInfo emailInfo = new EmailInfo();
                        emailInfo.NodeID = long.Parse(GetValue(data.data.properties, "id"));
                        emailInfo.ParentNodeID = long.Parse(GetValue(data.data.properties, "parent_id"));
                        emailInfo.Name = GetValue(data.data.properties, "name");
                        emailInfo.FileSize = long.Parse(GetValue(data.data.properties, "size"));
                        emailInfo.Summary = GetValue(data.data.properties, "summary");

                        emailInfo.EmailFrom = GetValue(data.data.regions, "OTEmailFrom");
                        emailInfo.EmailTo = GetValue(data.data.regions, "OTEmailTo");
                        emailInfo.EmailSubject = GetValue(data.data.regions, "OTEmailSubject");
                        emailInfo.HasAttachments = Convert.ToInt32(bool.Parse(GetValue(data.data.regions, "OTEmailHasAttachments")));

                        result.EmailInfos.Add(emailInfo);
                    }
                    result.Remarks = $"Search Results: {searchResponse.collection.paging.range_min} to {searchResponse.collection.paging.range_max} of about {searchResponse.collection.paging.total_count}";
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Something went wrong when trying to parse the Search Response: {ex.Message}\r\n{ex.StackTrace}");
            }

            return result;
        }

        private static string GetValue(Dictionary<string, object> dict, string key)
        {
            string valueStr = "";

            try
            {
                if (dict.ContainsKey(key))
                {
                    var value = dict[key];

                    if (value != null)
                    {
                        valueStr = dict[key].ToString();
                    }
                    else
                    {
                        logger.Debug($"Key: {key} value is null");
                    }
                }
                else
                {
                    logger.Debug($"Key: {key} does not exist");
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Something went wrong when attempting to get the value from {key}: {ex.Message}\r\n{ex.StackTrace}");
            }

            return valueStr;
        }
    }
}