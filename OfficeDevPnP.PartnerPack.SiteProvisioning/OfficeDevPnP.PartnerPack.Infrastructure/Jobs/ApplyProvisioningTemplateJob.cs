using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    public class ApplyProvisioningTemplateJob : ProvisioningJob
    {
        /// <summary>
        /// Defines the Provisioning Template to use for the Site Collection or Sub site to provision
        /// </summary>
        public String ProvisioningTemplateUrl { get; set; }

        /// <summary>
        /// Defines the Target Site URL of the Provisioning Template
        /// </summary>
        public String TargetSiteUrl { get; set; }
    }
}
