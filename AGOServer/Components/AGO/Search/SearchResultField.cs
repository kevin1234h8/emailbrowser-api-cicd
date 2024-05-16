using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer
{
    public class SearchResultField
    {
        public string Value { get; set; }
        private SearchResultField(string value)
        {
            Value = value;
        }

        //Score is not allowed to be set as Result Transformation Spec if we use SortByRegion
        //public static SearchResultField Score { get { return new SearchResultField("Score"); } }
        public static SearchResultField OTName { get { return new SearchResultField("OTName"); } }
        public static SearchResultField OTLocation { get { return new SearchResultField("OTLocation"); } }


        public static SearchResultField To { get { return new SearchResultField("OTEmailTo"); } }
        public static SearchResultField From { get { return new SearchResultField("OTEmailFrom"); } }
        public static SearchResultField CC { get { return new SearchResultField("OTEmailCC"); } }
        public static SearchResultField Subject { get { return new SearchResultField("OTEmailSubject"); } }
        public static SearchResultField HasAttachments { get { return new SearchResultField("OTEmailHasAttachments"); } }
        public static SearchResultField ReceivedDate { get { return new SearchResultField("OTEmailReceivedDate"); } }
        public static SearchResultField SentDate { get { return new SearchResultField("OTEmailSentDate"); } }
        public static SearchResultField FileSize { get { return new SearchResultField("OTDataSize"); } }
        
        public static SearchResultField Summary { get { return new SearchResultField("OTSummary"); } }



        public static SearchResultField Description { get { return new SearchResultField("Description"); } }
    }
}