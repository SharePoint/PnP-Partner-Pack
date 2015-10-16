using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines a Sub Site Provisioning Job
    /// </summary>
    public class SubSiteProvisioningJob : SiteProvisioningJob
    {
        /// <summary>
        /// Declares whether to inherit permissions from the parent Site Collection during Site provisioning
        /// </summary>
        public Boolean InheritPermissions { get; set; }
    }
}
