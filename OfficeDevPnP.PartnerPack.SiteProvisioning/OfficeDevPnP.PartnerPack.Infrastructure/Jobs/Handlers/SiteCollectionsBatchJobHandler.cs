using Microsoft.IdentityModel.Claims;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    public class SiteCollectionsBatchJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            SiteCollectionsBatchJob batchJob = job as SiteCollectionsBatchJob;
            if (batchJob == null)
            {
                throw new ArgumentException("Invalid job type for SiteCollectionsBatchJobHandler.");
            }

            CreateSiteCollectionsBatch(batchJob);
        }

        private void CreateSiteCollectionsBatch(SiteCollectionsBatchJob job)
        {
            var batches = JsonConvert.DeserializeObject<batches>(job.BatchSites);

            // For each Site Collection that we have to create
            foreach (var batch in batches.siteCollection)
            {
                // Prepare the Job to provision the Site Collection
                SiteCollectionProvisioningJob siteJob = new SiteCollectionProvisioningJob();

                // Prepare all the other information about the Provisioning Job
                siteJob.SiteTitle = batch.title;
                siteJob.Description = batch.description;
                siteJob.Language = Int32.Parse(batch.language);
                siteJob.TimeZone = batch.timeZone;
                siteJob.RelativeUrl = String.Format("/{0}/{1}", batch.managedPath, batch.relativeUrl);
                siteJob.SitePolicy = batch.sitePolicy == baseSiteSettingsSitePolicy.LBI ? "LBI" :
                    batch.sitePolicy == baseSiteSettingsSitePolicy.MBI ? "MBI" :
                    "HBI";
                siteJob.Owner = job.Owner;
                siteJob.PrimarySiteCollectionAdmin = batch.primarySiteCollectionAdmin;
                siteJob.SecondarySiteCollectionAdmin = batch.secondarySiteCollectionAdmin;
                siteJob.ProvisioningTemplateUrl = batch.templateUrl;
                siteJob.TemplatesProviderTypeName = batch.templatesProviderName;
                siteJob.StorageMaximumLevel = batch.storageMaximulLevel;
                siteJob.StorageWarningLevel = batch.storageWarningLevel;
                siteJob.UserCodeMaximumLevel = 0;
                siteJob.UserCodeWarningLevel = 0;
                siteJob.ExternalSharingEnabled = batch.externalSharingEnabled;
                siteJob.ResponsiveDesignEnabled = batch.responsiveDesignEnabled;
                siteJob.PartnerPackExtensionsEnabled = batch.partnerPackExtensionsEnabled;
                siteJob.ApplyTenantBranding = batch.applyTenantBranding;
                siteJob.Title = String.Format("Provisioning of Site Collection \"{1}\" with Template \"{0}\" by {2}",
                    siteJob.ProvisioningTemplateUrl,
                    siteJob.RelativeUrl,
                    siteJob.Owner);

                if (batch.templateParameters != null)
                {
                    siteJob.TemplateParameters = batch.templateParameters.ToDictionary(i => i.Key, i => i.Value);
                }

                var jobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(siteJob);

                Console.WriteLine($"Scheduled Site Collection creation job with ID: {jobId}");
            }
        }
    }
}
