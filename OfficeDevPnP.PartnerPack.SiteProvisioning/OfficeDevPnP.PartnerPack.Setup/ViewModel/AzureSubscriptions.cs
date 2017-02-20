using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.ViewModel
{
    public class AzureSubscriptions
    {
        [JsonProperty(PropertyName = "Value")]
        public AzureSubscription[] Subscriptions { get; set; }
    }
}
