using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class AzureAdApplicationPasswordCredential
    {
        public object CustomKeyIdentifier { get; set; }

        [JsonProperty("startDateTime")]
        public string StartDate { get; set; }
        [JsonProperty("endDateTime")]
        public string EndDate { get; set; }

        public string KeyId { get; set; }

        [JsonProperty("secretText")]
        public object Value { get; set; }
    }
}
