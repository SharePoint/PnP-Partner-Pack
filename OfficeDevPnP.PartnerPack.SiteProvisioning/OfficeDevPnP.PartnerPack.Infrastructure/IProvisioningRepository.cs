using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Interface that defines the common behavior for any Sites Provisioning Repository
    /// </summary>
    public interface IProvisioningRepository : IConfigurable
    {
        /// <summary>
        /// Retrieves the list of Global Provisioning Templates
        /// </summary>
        /// <param name="scope">The scope to filter the provisioning templates</param>
        /// <returns>Returns the list of Provisioning Templates</returns>
        ProvisioningTemplateInformation[] GetGlobalProvisioningTemplates(TargetScope scope);

        /// <summary>
        /// Retrieves the list of Local Provisioning Templates
        /// </summary>
        /// <param name="siteUrl">The local Site Collection to retrieve the templates from</param>
        /// <param name="scope">The scope to filter the provisioning templates</param>
        /// <returns>Returns the list of Provisioning Templates</returns>
        ProvisioningTemplateInformation[] GetLocalProvisioningTemplates(String siteUrl, TargetScope scope);

        /// <summary>
        /// Saves a Provisioning Template into the target Global repository
        /// </summary>
        /// <param name="template">The Provisioning Template to save</param>
        void SaveGlobalProvisioningTemplate(GetProvisioningTemplateJob job);

        /// <summary>
        /// Saves a Provisioning Template into the target Local repository
        /// </summary>
        /// <param name="siteUrl">The local Site Collection to save to</param>
        /// <param name="template">The Provisioning Template to save</param>
        void SaveLocalProvisioningTemplate(String siteUrl, GetProvisioningTemplateJob job);

        /// <summary>
        /// Enqueues a new Provisioning Job
        /// </summary>
        /// <param name="job">The Provisioning Job to enqueue</param>
        /// <returns>Returns the ID of the job</returns>
        Guid EnqueueProvisioningJob(ProvisioningJob job);

        /// <summary>
        /// Updates a job in the queue
        /// </summary>
        /// <remarks>In case of failure it will throw an Exception</remarks>
        /// <param name="job">The information about the job to update</param>
        void UpdateProvisioningJob(Guid jobId, ProvisioningJobStatus status, String errorMessage = null);

        /// <summary>
        /// Retrieves the list of Provisioning Jobs
        /// </summary>
        /// <param name="status">The status to use for filtering Provisioning Jobs</param>
        /// <param name="includeStream">Defines whether to include the stream of the serialized job</param>
        /// <param name="owner">The optional owner of the Provisioning Job</param>
        /// <returns>The list of information about the Provisioning Jobs, if any</returns>
        ProvisioningJobInformation[] GetProvisioningJobs(ProvisioningJobStatus status, String jobType = null, Boolean includeStream = false, String owner = null);

        /// <summary>
        /// Retrieves a Provisioning Job by ID
        /// </summary>
        /// <param name="jobId">The ID of the job to retrieve</param>
        /// <param name="includeStream">Defines whether to include the stream of the serialized job</param>
        /// <returns>The information about the Provisioning Job, if any</returns>
        ProvisioningJobInformation GetProvisioningJob(Guid jobId, Boolean includeStream = false);

        /// <summary>
        /// Retrieves the list of Provisioning Jobs
        /// </summary>
        /// <param name="status">The status to use for filtering Provisioning Jobs</param>
        /// <param name="owner">The optional owner of the Provisioning Job</param>
        /// <typeparam name="TJob">Represents the type of the Provisioning Jobs to retrieve</typeparam>
        /// <returns>The list of information about the Provisioning Jobs, if any</returns>
        ProvisioningJob[] GetTypedProvisioningJobs<TJob>(ProvisioningJobStatus status, String owner = null)
            where TJob : ProvisioningJob;
    }
}
