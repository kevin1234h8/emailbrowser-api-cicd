using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace AGOServer
{

    public class EmailSearchRequest
    {
        string keywords;
        long folderNodeID = 2000;
        int firstResultToRetrieve = 1;
        int numResultsToRetrieve = 200;

        /// <summary>
        /// The keywords of the search, does not have to be url encoded for now. supports spaces currently.
        /// </summary>
        [Required]
        public string Keywords { get => keywords; set => keywords = value; }

        /// <summary>
        /// Search in this folder id, recursively. Default is 2000 for enterprise
        /// </summary>
        public long FolderNodeID { get => folderNodeID; set => folderNodeID = value; }
        /// <summary>
        /// The start index of the result, default is 1. For example, 
        /// the total number of results is 500, and you have shown results 1 to 200 (with include count being 200), 
        /// you would want the FirstResultToRetrieve to be 201.
        /// </summary>
        public int FirstResultToRetrieve { get => firstResultToRetrieve; set => firstResultToRetrieve = value; }

        /// <summary>
        /// The page size, how many results do you want to retrieve. Default is 200
        /// </summary>
        public int NumResultsToRetrieve { get => numResultsToRetrieve; set => numResultsToRetrieve = value; }
    }
}