using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SharePointCamlQuery.Models
{
    public class ListQueryResponse
    {
        [JsonProperty("d")]
        public ListQueryResults d { get; set; }
    }
    public class ListQueryResults
    {
        [JsonProperty("results")]
        public List<ListInfo> Results { get; set; }
    }
    public class ListInfo
    {
        public int BaseTemplate { get; set; }
        public int BaseType { get; set; }
        public bool ContentTypesEnabled { get; set; }
        public bool CrawlNonDefaultViews { get; set; }
        public DateTime Created { get; set; }
        public string Description { get; set; }
        public string EntityTypeName { get; set; }
        public bool Hidden { get; set; }
        public string Id { get; set; }
        public bool IsApplicationList { get; set; }
        public bool IsCatalog { get; set; }
        public bool IsPrivate { get; set; }
        public int ItemCount { get; set; }
        public DateTime LastItemDeletedDate { get; set; }
        public DateTime LastItemModifiedDate { get; set; }
        public string Title { get; set; }
    }
}
