using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace testHubSpot
{
    public partial class OriginContact
    {
        [JsonProperty("vid")]
        public long Vid { get; set; }


        [JsonProperty("properties")]
        public Dictionary<string, ContactProperty> Properties { get; set; }

        [JsonProperty("associated-company")]
        public AssociatedCompany AssociatedCompany { get; set; }
    }

    public partial class AssociatedCompany
    {
        [JsonProperty("company-id")]
        public long CompanyId { get; set; }

        [JsonProperty("properties")]
        public Dictionary<string, AssociatedCompanyProperty> Properties { get; set; }
    }

    public partial class AssociatedCompanyProperty
    {
        [JsonProperty("value")]
        public string Value { get; set; }
    }

    public partial class ContactProperty
    {
        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("versions")]
        public List<Version> Versions { get; set; }
    }
}
