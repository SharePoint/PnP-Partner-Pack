using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.SharePoint
{
    /// <summary>
    /// Provides the implementation of a Provisioning Repository that targets SharePoint
    /// </summary>
    public class SharePointProvisioningRepository : IProvisioningRepository
    {
        public void Init()
        {
            // NOOP
            return;
        }

        public ProvisioningTemplateInformation[] GetGlobalProvisioningTemplates(TemplateScope scope)
        {
            throw new NotImplementedException();
        }

        public ProvisioningTemplateInformation[] GetLocalProvisioningTemplates(string siteUrl, TemplateScope scope)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Saves a Provisioning Template into the target Global repository
        /// </summary>
        /// <param name="job">The Provisioning Template to save</param>
        /// <returns>The ID  of the saved Provisioning Template</returns>
        public Guid SaveGlobalProvisioningTemplate(GetProvisioningTemplateJob job)
        {
            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                return (SaveProvisioningTemplateInternal(context, job));
            }
        }

        /// <summary>
        /// Saves a Provisioning Template into the target Local repository
        /// </summary>
        /// <param name="siteUrl">The local Site Collection to save to</param>
        /// <param name="template">The Provisioning Template to save</param>
        /// <returns>The ID  of the saved Provisioning Template</returns>
        public Guid SaveLocalProvisioningTemplate(string siteUrl, GetProvisioningTemplateJob job)
        {
            // Connect to the Local Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                PnPPartnerPackUtilities.EnablePartnerPackInfrastructureOnSite(siteUrl);
                return (SaveProvisioningTemplateInternal(context, job));
            }
        }

        private Guid SaveProvisioningTemplateInternal(ClientContext context, GetProvisioningTemplateJob job)
        {
            // Get a reference to the target web site
            Web web = context.Web;
            context.Load(web, w => w.Url);
            context.ExecuteQueryRetry();

            // Configure the XML file system provider
            XMLTemplateProvider provider =
                new XMLSharePointTemplateProvider(context, web.Url, 
                    PnPPartnerPackConstants.PnPProvisioningTemplates);

            ProvisioningTemplateCreationInformation ptci =
                new ProvisioningTemplateCreationInformation(web);
            ptci.FileConnector = provider.Connector;
            ptci.IncludeAllTermGroups = job.IncludeAllTermGroups;
            ptci.IncludeSearchConfiguration = job.IncludeSearchConfiguration;
            ptci.IncludeSiteCollectionTermGroup = job.IncludeSiteCollectionTermGroup;
            ptci.IncludeSiteGroups = job.IncludeSiteGroups;
            ptci.PersistComposedLookFiles = job.PersistComposedLookFiles;

            // Extract the current template
            ProvisioningTemplate templateToSave = web.GetProvisioningTemplate(ptci);

            templateToSave.Description = job.Description;
            templateToSave.DisplayName = job.Title;

            // TODO: Implement this one
            templateToSave.ImagePreviewUrl = "fake.png"; 

            // And save it on the file system
            provider.SaveAs(templateToSave, job.FileName);

            // TODO: Replace with real ID of the file
            return (Guid.NewGuid());
        }

        public Guid EnqueueProvisioningJob(ProvisioningJob job)
        {
            // Prepare the Job ID
            Guid jobId = Guid.NewGuid();

            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Convert the current Provisioning Job into a Stream
                Stream stream = job.ToJsonStream();

                // Get a reference to the target library
                Web web = context.Web;
                List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningJobs);
                Microsoft.SharePoint.Client.File file = list.RootFolder.UploadFile(String.Format("{0}.job", jobId), stream, false);

                ListItem item = file.ListItemAllFields;
                item[PnPPartnerPackConstants.ContentTypeIdField] = PnPPartnerPackConstants.PnPProvisioningJobContentTypeId;
                item[PnPPartnerPackConstants.TitleField] = job.Title;
                item[PnPPartnerPackConstants.PnPProvisioningJobStatus] = ProvisioningJobStatus.Pending.ToString();
                item[PnPPartnerPackConstants.PnPProvisioningJobError] = String.Empty;
                item[PnPPartnerPackConstants.PnPProvisioningJobType] = job.GetType().FullName;
                item.Update();

                context.ExecuteQueryRetry();
            }

            return (jobId);
        }

        public ProvisioningJob GetProvisioningJob(Guid jobId)
        {
            throw new NotImplementedException();
        }

        public ProvisioningJob[] GetProvisioningJobs(ProvisioningJobStatus status, string owner = null)
        {
            throw new NotImplementedException();
        }

        public void UpdateProvisioningJob(ProvisioningJob job)
        {
            throw new NotImplementedException();
        }
    }
}
