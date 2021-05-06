using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SharePointCamlQuery.Models
{
    public class Metadata
    {
        public string id { get; set; }
        public string uri { get; set; }
        public string etag { get; set; }
        public string type { get; set; }
    }

    public class DocumentModel
    {
        public Metadata __metadata { get; set; }
        public int FileSystemObjectType { get; set; }
        [JsonProperty("Id")]
        public int IdLowerCase { get; set; }
        [DisplayFormat(DataFormatString = "{0:MMM d, yyyy h:mm tt }")]
        public DateTime Created { get; set; }
        [JsonProperty("ID")]
        public int ID { get; set; }
        public string FileLeafRef { get; set; }
        public string EncodedAbsUrl { get; set; }
        [DisplayFormat(DataFormatString = "{0:MMM d, yyyy h:mm tt }")]
        public DateTime Modified { get; set; }
    }

    public class DocumentQueryResults
    {
        [JsonProperty("results")]
        public List<DocumentModel> Results { get; set; }
    }

    public class DocumentQueryResponse
    {
        [JsonProperty("d")]
        public DocumentQueryResults d { get; set; }
    }
}
