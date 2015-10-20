using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.TimerJobs;
using OfficeDevPnP.PartnerPack.Infrastructure;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.ProvisioningJob
{
    public class PnPPartnerPackProvisioningJob : TimerJob
    {
        public PnPPartnerPackProvisioningJob(): base("PnP Partner Pack Provisioning Job")
        {
            TimerJobRun += ExecuteProvisioningJobs;
        }
        private void ExecuteProvisioningJobs(object sender, TimerJobRunEventArgs e)
        {
            Console.WriteLine("Starting job");
            Web web = e.SiteClientContext.Web;
            web.EnsureProperty(w => w.Title);
            Console.WriteLine("Processing jobs in Site: {0}", web.Title);

            // Retrieve the list of pending jobs
            var provisioningJobs = ProvisioningRepositoryFactory.Current.GetProvisioningJobs(ProvisioningJobStatus.Pending);
            foreach (var pj in provisioningJobs)
            {
                Console.WriteLine("Processing job: {0} - Owner: {1} - Title: {2}", 
                    pj.JobId, pj.Owner, pj.Title);

                // Deserialize the job
                var job = pj.JobFile.FromJsonStream(pj.Type);

                // Process the job
                if (job  is ApplyProvisioningTemplateJob)
                {
                    // Provisioning Template Application
                }
                else if (job is GetProvisioningTemplateJob)
                {
                    // Get Provisioning Template
                    GetProvisioningTemplateJob gptj = job as GetProvisioningTemplateJob;
                    if (gptj.Location == ProvisioningTemplateLocation.Global)
                    {
                        ProvisioningRepositoryFactory.Current.SaveGlobalProvisioningTemplate(gptj);
                    }
                    else
                    {
                        ProvisioningRepositoryFactory.Current.SaveLocalProvisioningTemplate(gptj.SourceSiteUrl, gptj);
                    }
                }
                else if (job is SiteCollectionProvisioningJob)
                {
                    // Site Collection Provisioning
                }
                else if (job is SubSiteProvisioningJob)
                {
                    // Sub-Site Provisioning
                }
            }

            Console.WriteLine("Ending job");
        }
    }
}
