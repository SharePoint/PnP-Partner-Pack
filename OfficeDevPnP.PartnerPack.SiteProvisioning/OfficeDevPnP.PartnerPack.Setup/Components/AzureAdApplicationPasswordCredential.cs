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
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string KeyId { get; set; }
        public object Value { get; set; }
    }
}
