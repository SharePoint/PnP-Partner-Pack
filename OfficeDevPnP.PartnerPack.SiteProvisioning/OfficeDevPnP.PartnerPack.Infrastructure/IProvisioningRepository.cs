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
    public interface IProvisioningRepository
    {
        /// <summary>
        /// Enqueues a new Site Collection creation job
        /// </summary>
        /// <param name="site">The information about the Site Collection to create</param>
        /// <returns>Returns the ID of the job</returns>
        Guid EnqueueSiteCollectionCreation(SiteCollectionCreationInformation site);

        /// <summary>
        /// Enqueues a new Sub Site creation job
        /// </summary>
        /// <param name="site">The information about the Sub Site to create</param>
        /// <returns>Returns the ID of the job</returns>
        Guid EnqueueSubSiteCreation(SubSiteCreationInformation site);

        /// <summary>
        /// Enqueues a job to provision a template onto a target Site
        /// </summary>
        /// <param name="template">The information about the template to apply</param>
        /// <returns>Returns the ID of the job</returns>
        Guid EnqueueProvisioningTemplateApplication(ProvisioningTemplateApplicationInformation template);

        /// <summary>
        /// Retrieves the list of Global Provisioning Templates
        /// </summary>
        /// <param name="scope">The scope to filter the provisioning templates</param>
        /// <returns>Returns the list of Provisioning Templates</returns>
        ProvisioningTemplateInformation[] GetGlobalProvisioningTemplates(TemplateScope scope);

        /// <summary>
        /// Retrieves the list of Local Provisioning Templates
        /// </summary>
        /// <param name="siteUrl">The local Site Collection to retrieve the templates from</param>
        /// <param name="scope">The scope to filter the provisioning templates</param>
        /// <returns>Returns the list of Provisioning Templates</returns>
        ProvisioningTemplateInformation[] GetLocalProvisioningTemplates(String siteUrl, TemplateScope scope);

        /// <summary>
        /// Updates a job in the queue
        /// </summary>
        /// <remarks>In case of failure it will throw an Exception</remarks>
        /// <param name="job">The job to update</param>
        void UpdateJob(ProvisioningJob job);
        
        /// <summary>
        /// Retrieves the list of Provisioning Jobs
        /// </summary>
        /// <param name="status">The status to use for filtering Provisioning Jobs</param>
        /// <param name="owner">The optional owner of the Provisioning Job</param>
        /// <returns>The list of Provisioning Jobs</returns>
        ProvisioningJob[] GetJobs(JobStatus status, String owner = null);

        /// <summary>
        /// Retrieves a Provisioning Jobs by ID
        /// </summary>
        /// <param name="jobId">The ID of the job to retrieve</param>
        /// <returns>The Provisioning Jobs, if any</returns>
        ProvisioningJob GetJob(Guid jobId);
    }
}
