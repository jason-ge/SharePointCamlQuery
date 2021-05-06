using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SharePointCamlQuery.Models
{
    public class FieldInfo
    {
        public string Title { get; set; }
        public string InternalName { get; set; }
    }

    public class FieldInfoResults
    {
        [JsonProperty("results")]
        public List<FieldInfo> Results { get; set; }
    }

    public class FieldInfoResponse
    {
        [JsonProperty("d")]
        public FieldInfoResults d { get; set; }
    }
}
