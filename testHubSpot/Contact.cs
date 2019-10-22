using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testHubSpot
{
    class Contact
    {
        public long Vid { get; set; }
        public string FirstName { get; set; } = "empty";
        public string SecondName { get; set; } = "empty";
        public string Lifecyclestage { get; set; } = "empty";
        public long Company_id { get; set; } = 0;
        public string CompanyName { get; set; } = "empty";
        public string Website { get; set; } = "www.___.com";
        public string City { get; set; } = "empty";
        public string State { get; set; } = "empty";
        public string Zip { get; set; } = "empty";
        public string Phone { get; set; } = "+##-###-##-##-###";
    }
}
