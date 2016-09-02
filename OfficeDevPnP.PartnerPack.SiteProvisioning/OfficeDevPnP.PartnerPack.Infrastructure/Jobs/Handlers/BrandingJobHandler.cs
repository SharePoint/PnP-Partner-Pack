using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    public class BrandingJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            // For each Site Collection in the tenant
            // Apply branding (if it is not already applied)
            // Store Property Bag to set the last version of branding applied
        }
    }
}
