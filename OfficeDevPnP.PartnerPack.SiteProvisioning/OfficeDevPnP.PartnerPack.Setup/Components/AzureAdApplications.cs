using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class AzureAdApplications
    {
        [JsonProperty("value")]
        public List<AzureAdApplication> Applications { get; set; }
    }
}
