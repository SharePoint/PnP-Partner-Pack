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
using System.Diagnostics;
using OfficeDevPnP.PartnerPack.Infrastructure.Diagnostics;

namespace OfficeDevPnP.PartnerPack.ContinousJob
{
    public class Functions
    {
        // This function will get triggered/executed when a new message is written 
        // on an Azure Queue called queue.
        public static void ProcessQueueMessage([QueueTrigger(PnPPartnerPackSettings.StorageQueueName)] ContinousJobItem content, TextWriter log)
        {
            // Attach TextWriter log to Trace
            // https://blog.josequinto.com/2017/02/16/enable-azure-invocation-log-at-a-web-job-function-level-for-pnp-provisioning/
            TextWriterTraceListener twtl = new TextWriterTraceListener(log);
            twtl.Name = "ContinousJobLogger";
            string[] notShownWords = new string[] { "TokenCache", "AcquireTokenHandlerBase"};
            twtl.Filter = new RemoveWordsFilter(notShownWords); 
            Trace.Listeners.Add(twtl);
            Trace.AutoFlush = true;

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

            // Remove Trace Listener
            Trace.Listeners.Remove(twtl);
        }
    }
}
