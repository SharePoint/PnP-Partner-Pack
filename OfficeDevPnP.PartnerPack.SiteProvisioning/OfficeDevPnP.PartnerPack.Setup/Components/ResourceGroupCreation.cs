using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class Properties
    {
        public string provisioningState { get; set; }
    }

    public class ResourceGroupCreation
    {
        public string id { get; set; }
        public string name { get; set; }
        public string location { get; set; }
        public Properties properties { get; set; }
    }
}
