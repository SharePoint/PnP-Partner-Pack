using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    /// <summary>
    /// Abstract type for any Provisioning Job Handler
    /// </summary>
    public abstract class ProvisioningJobHandler : IConfigurable
    {
        /// <summary>
        /// Initialization method, if needed
        /// </summary>
        /// <param name="configuration">The Provisioning Job Handler configuration</param>
        public virtual void Init(XElement configuration = null)
        {
            // NOOP;
            return;
        }

        /// <summary>
        /// Internal implementation for running a Provisioning Job
        /// </summary>
        /// <param name="job">The Provisioning Job to run</param>
        protected abstract void RunJobInternal(ProvisioningJob job);

        /// <summary>
        /// Executes a Provisioning Job
        /// </summary>
        /// <param name="job">The Provisioning Job to run</param>
        public void RunJob(ProvisioningJob job)
        {
            try
            {
                // Set the Job status as Running
                ProvisioningRepositoryFactory.Current.UpdateProvisioningJob(
                job.JobId,
                ProvisioningJobStatus.Running,
                String.Empty);

                // Run the Job
                RunJobInternal(job);

                // Set the Job status as Provisioned (i.e. Completed)
                ProvisioningRepositoryFactory.Current.UpdateProvisioningJob(
                    job.JobId,
                    ProvisioningJobStatus.Provisioned,
                    String.Empty);
            }
            catch (Exception ex)
            {
                // Set the Job status as Failed, including the exception details
                ProvisioningRepositoryFactory.Current.UpdateProvisioningJob(
                    job.JobId,
                    ProvisioningJobStatus.Failed,
                    ex.Message);

                Console.WriteLine("Exception occurred: {0}\nStack Trace:\n{1}\n", ex.Message, ex.StackTrace);
            }
        }
    }
}
