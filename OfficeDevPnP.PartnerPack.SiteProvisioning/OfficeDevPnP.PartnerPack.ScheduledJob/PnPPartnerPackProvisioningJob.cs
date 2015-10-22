using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.TimerJobs;
using OfficeDevPnP.PartnerPack.Infrastructure;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.ScheduledJob
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

            // Show the current context
            Web web = e.SiteClientContext.Web;
            web.EnsureProperty(w => w.Title);
            Console.WriteLine("Processing jobs in Site: {0}", web.Title);

            // Retrieve the list of pending jobs
            var provisioningJobs = ProvisioningRepositoryFactory.Current.GetTypedProvisioningJobs<ProvisioningJob>(
                ProvisioningJobStatus.Pending);

            foreach (var job in provisioningJobs)
            {
                Console.WriteLine("Processing job: {0} - Owner: {1} - Title: {2}",
                    job.JobId, job.Owner, job.Title);

                Type jobType = job.GetType();

                if (PnPPartnerPackSettings.JobHandlers.ContainsKey(jobType))
                {
                    PnPPartnerPackSettings.JobHandlers[jobType].RunJob(job);
                }
            }

            Console.WriteLine("Ending job");
        }
    }
}
