using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    public class SiteCollectionProvisioningJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            SiteCollectionProvisioningJob scj = job as SiteCollectionProvisioningJob;
            if (scj == null)
            {
                throw new ArgumentException("Invalid job type for SiteCollectionProvisioningJobHandler.");
            }

            return;
        }
    }
}
