using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines an Refresh Job for a single site
    /// </summary>
    public class RefreshSingleSiteJob : ProvisioningJob
    {
        /// <summary>
        /// Defines the Target Site URL of the Provisioning Template
        /// </summary>
        public String TargetSiteUrl { get; set; }
    }
}
