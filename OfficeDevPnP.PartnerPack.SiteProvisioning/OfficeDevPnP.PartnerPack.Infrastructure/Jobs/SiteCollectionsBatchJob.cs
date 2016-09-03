using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines a Site Collections Batch Provisioning Job
    /// </summary>
    public class SiteCollectionsBatchJob : ProvisioningJob
    {
        /// <summary>
        /// Defines a JSON representation of the sites to create
        /// </summary>
        public String BatchSites { get; set; }
    }
}
