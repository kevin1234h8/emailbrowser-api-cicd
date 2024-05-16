using AGOServer;
using CSAccess_Local;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;

namespace AGOServerDeploymentTests
{
    public class AGOServerDeploymentTests
    {
        static long folderID = Properties.Settings.Default.FolderThatContainsEmail;
        static string username = Properties.Settings.Default.ImpersonationUserNameForTesting;
        public static void AllTests()
        {
            //SearchForEmails_Test();
            //ListAttachments_Test();
            ListEmails_Test();
            //FindConversationEmails_Test();
            //ListPaginatedEmails_Test();
            //ListNodes_REST_Test();
            Console.ReadKey();
        }
        static void PrintHeader(string header)
        {
            Console.WriteLine(System.Environment.NewLine);
            Console.WriteLine("########################## " + header + " #############################");
        }
        static void PrintObject(Object obj)
        {
            var jsonString = JsonConvert.SerializeObject(
           obj, Formatting.Indented,
           new JsonConverter[] { new StringEnumConverter() });
            Console.WriteLine(jsonString);
        }
        static void ListEmails_Test()
        {
            PrintHeader("ListNodes_Test");
            var result = ListingService.ListEmails(username, folderID, true, "OTEmailTo", "desc");
            PrintObject(result);
        }
        static void ListPaginatedEmails_Test()
        {
            PrintHeader("ListNodes_Test");
            //149666
            var result = ListingService.ListEmails(username, folderID, false, 1, 50, "OTEmailTo", "desc");
            PrintObject(result);
        }
        static void FindConversationEmails_Test()
        {
            PrintHeader("FindConversationEmails_Test");
            long emailId = Properties.Settings.Default.EmailNodeID_WithConversations;
            var result = ListingService.FindConversationEmails(username, emailId, "OTEmailSentDate", "desc");
            PrintObject(result);
        }

        static void ListAttachments_Test()
        {
            PrintHeader("ListAttachments");
            long emailId = Properties.Settings.Default.EmailNodeID_WithAttachment;
            var result = AttachmentServices.ListAttachments(username, emailId);
            PrintObject(result);
        }
        static void SearchForEmails_Test()
        {
            PrintHeader("SearchForEmails");
            string OTEmailReceivedDate_After = null;
            string OTEmailReceivedDate_Before = null;
            SearchFilterByAttachment searchFilterByAttachment = SearchFilterByAttachment.UNDEFINED;
            string OTEmailSubject = null;
            string OTEmailFrom = null;
            string OTEmailTo = "";
            bool includeConversationID = true;
            string attachmentFileName = "";

            //var result = EmailSearchService.SearchForEmails(username, "", folderID, 1, 100, "OTEmailSubject", "asc",
            // OTEmailReceivedDate_After, OTEmailReceivedDate_Before, searchFilterByAttachment,
            // OTEmailSubject, OTEmailFrom, OTEmailTo, includeConversationID, attachmentFileName);
            //PrintObject(result);
            var result = EmailSearchService.SearchForEmails(username, "", folderID, 1, 100, "OTEmailSubject", "asc",
     OTEmailReceivedDate_After, OTEmailReceivedDate_Before, searchFilterByAttachment,
     OTEmailSubject, OTEmailFrom, OTEmailTo, includeConversationID, attachmentFileName);
            PrintObject(result);


        }
    }
}