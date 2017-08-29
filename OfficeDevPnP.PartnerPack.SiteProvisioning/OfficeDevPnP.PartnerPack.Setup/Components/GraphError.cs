using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class InnerError
    {
        [JsonProperty("request-id")]
        public Guid requestId { get; set; }
        public string date { get; set; }
    }

    public class Error
    {
        public string code { get; set; }
        public string message { get; set; }
        public InnerError innerError { get; set; }
    }

    public class GraphError
    {
        public Error error { get; set; }
    }
}
