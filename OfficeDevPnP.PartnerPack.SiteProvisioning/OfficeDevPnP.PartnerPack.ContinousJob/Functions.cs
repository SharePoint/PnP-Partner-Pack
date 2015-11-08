using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers;
using OfficeDevPnP.PartnerPack.Infrastructure;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;

namespace OfficeDevPnP.PartnerPack.ContinousJob
{
    public class Functions
    {
        // This function will get triggered/executed when a new message is written 
        // on an Azure Queue called queue.
        public static void ProcessQueueMessage([QueueTrigger(PnPPartnerPackSettings.StorageQueueName)] ContinousJobItem content, TextWriter log)
        {
            log.WriteLine(String.Format("Found Job: {0}", content.JobId));

            // Get the info about the Provisioning Job
            ProvisioningJobInformation jobInfo =
                ProvisioningRepositoryFactory.Current.GetProvisioningJob(content.JobId, true);

            // Get a reference to the Provisioning Job
            ProvisioningJob job = jobInfo.JobFile.FromJsonStream(jobInfo.Type);

            if (PnPPartnerPackSettings.ContinousJobHandlers.ContainsKey(job.GetType()))
            {
                PnPPartnerPackSettings.ContinousJobHandlers[job.GetType()].RunJob(job);
            }

            log.WriteLine("Completed Job execution");
        }
    }
}
