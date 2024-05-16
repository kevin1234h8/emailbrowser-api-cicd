using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AGOServer.Components.Models.OpenText
{
    /// <summary>
    /// This Data Model is designed based on the schema available via OT Dev Portal
    /// </summary>
    public class OpenTextV2SearchResponse
    {
        public search_collectionData collection { get; set; }
        public search_linksData links { get; set; }
        public List<search_resultsData> results { get; set; }
    }

    public class search_collectionData
    {
        public search_pagingData paging { get; set; }
        public search_searchingData searching { get; set; }
        public search_sortingData sorting { get; set; }
    }

    public class search_pagingData
    {
        public int limit { get; set; }
        public int page { get; set; }
        public int page_total { get; set; }
        public int range_max { get; set; }
        public int range_min { get; set; }
        public string result_header_string { get; set; }
        public int total_count { get; set; }
    }

    public class search_searchingData
    {
        public int cache_id { get; set; }
        public object facets { get; set; }
        public string result_title { get; set; }
        public object regions_order { get; set; }
        public object regions_metadata { get; set; }
    }

    public class search_sortingData
    {
        public search_linksSortingData links { get; set; }
        public List<string> sort { get; set; }
    }

    public class search_linksSortingData { }

    public class search_linksData
    {
        public search_searchLinksData data { get; set; }
    }

    public class search_searchLinksData
    {
        public search_searchLinksSelfData self { get; set; }
    }

    public class search_searchLinksSelfData
    {
        public string body { get; set; }
        public string content_type { get; set; }
        public string href { get; set; }
        public string method { get; set; }
        public string name { get; set; }
    }

    public class search_resultsData
    {
        public search_dataResultsData data { get; set; }
        public object links { get; set; }
        public string metadata { get; set; }
        public object search_result_metadata { get; set; }
    }

    public class search_dataResultsData
    {
        public Dictionary<string, object> properties { get; set; }
        public Dictionary<string, object> regions { get; set; }
        public object versions { get; set; }
    }
}