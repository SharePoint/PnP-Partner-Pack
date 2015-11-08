using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    public class ProvisioningTemplateJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            // Determine the type of job to run
            if (job is GetProvisioningTemplateJob)
            {
                RunGetProvisioningTemplateJob(job as GetProvisioningTemplateJob);
            }
            else if (job is ApplyProvisioningTemplateJob)
            {
                RunApplyProvisioningTemplateJob(job as ApplyProvisioningTemplateJob);
            }
            else
            {
                throw new ArgumentException("Invalid job type for ProvisioningTemplateJobHandler.");
            }
        }

        private void RunGetProvisioningTemplateJob(GetProvisioningTemplateJob job)
        {
            if (job.Location == ProvisioningTemplateLocation.Global)
            {
                ProvisioningRepositoryFactory.Current.SaveGlobalProvisioningTemplate(job);
            }
            else
            {
                ProvisioningRepositoryFactory.Current.SaveLocalProvisioningTemplate(job.SourceSiteUrl, job);
            }
        }

        private void RunApplyProvisioningTemplateJob(ApplyProvisioningTemplateJob job)
        {
            // TODO: Still missing implementation, will come in V.2 of the PnP Partner Pack
        }
    }
}
