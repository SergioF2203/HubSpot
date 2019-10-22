using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace testHubSpot
{
    public partial class RecentContacts
    {
        [JsonProperty("contacts")]
        public List<ContactElement> Contacts { get; set; }

        [JsonProperty("has-more")]
        public bool HasMore { get; set; }

        [JsonProperty("vid-offset")]
        public long VidOffset { get; set; }

        [JsonProperty("time-offset")]
        public long TimeOffset { get; set; }
    }

    public partial class ContactElement
    {
        [JsonProperty("vid")]
        public long Vid { get; set; }

        [JsonProperty("properties")]
        public Properties Properties { get; set; }

    }

    public partial class Properties
    {
        [JsonProperty("lastmodifieddate")]
        public Company Lastmodifieddate { get; set; }

    }

    public partial class Company
    {
        [JsonProperty("value")]
        public string Value { get; set; }
    }
}


