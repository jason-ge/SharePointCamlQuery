using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SharePointCamlQuery.Models
{
    public class SharePointSettings
    {
        public string SPUserName { get; set; }
        public string SPPassword { get; set; }
        public string SPDomain { get; set; }
        public string SiteUrl { get; set; }
    }
}
