using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines where to store a provisioning template
    /// </summary>
    public enum ProvisioningTemplateLocation
    {
        /// <summary>
        /// Store the template in the Global infrastructural Site Collection 
        /// </summary>
        Global,
        /// <summary>
        /// Store the template in the Local Site Collection
        /// </summary>
        Local,
    }
}
