using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    public class SubSiteProvisioningJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            SubSiteProvisioningJob ssj = job as SubSiteProvisioningJob;
            if (ssj == null)
            {
                throw new ArgumentException("Invalid job type for SubSiteProvisioningJobHandler.");
            }

            CreateSubSite(ssj);
        }

        private void CreateSubSite(SubSiteProvisioningJob job)
        {
            // Determine the reference URLs and relative paths
            String subSiteUrl = job.RelativeUrl;
            String siteCollectionUrl = PnPPartnerPackUtilities.GetSiteCollectionRootUrl(job.ParentSiteUrl);
            String parentSiteUrl = job.ParentSiteUrl;

            Console.WriteLine("Creating Site \"{0}\" as child Site of \"{1}\".", subSiteUrl, parentSiteUrl);

            using (ClientContext context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(parentSiteUrl))
            {
                // Get a reference to the parent Web
                Web parentWeb = context.Web;

                // Create the new sub site as a new child Web
                WebCreationInformation newWeb = new WebCreationInformation();
                newWeb.Description = job.Description;
                newWeb.Language = job.Language;
                newWeb.Title = job.SiteTitle;
                newWeb.Url = subSiteUrl;
                newWeb.UseSamePermissionsAsParentSite = job.InheritPermissions;
                newWeb.WebTemplate = PnPPartnerPackSettings.DefaultSiteTemplate;

                Web web = parentWeb.Webs.Add(newWeb);
                context.ExecuteQueryRetry();

                Console.WriteLine("Site \"{0}\" created.", subSiteUrl);

                // Apply the Provisioning Template
                Console.WriteLine("Applying Provisioning Template \"{0}\" to site.", 
                    job.ProvisioningTemplateUrl);

                // Determine the reference URLs and file names
                String templatesSiteUrl = job.ProvisioningTemplateUrl.Substring(0,
                    job.ProvisioningTemplateUrl.IndexOf(PnPPartnerPackConstants.PnPProvisioningTemplates));
                String templateFileName = job.ProvisioningTemplateUrl.Substring(job.ProvisioningTemplateUrl.LastIndexOf("/") + 1);

                // Configure the XML file system provider
                XMLTemplateProvider provider =
                    new XMLSharePointTemplateProvider(context, templatesSiteUrl,
                        PnPPartnerPackConstants.PnPProvisioningTemplates);

                // Load the template from the XML stored copy
                ProvisioningTemplate template = provider.GetTemplate(templateFileName);
                template.Connector = provider.Connector;

                // We do intentionally remove taxonomies, which are not supported in the AppOnly Authorization model
                // For further details, see the PnP Partner Pack documentation 
                ProvisioningTemplateApplyingInformation ptai =
                    new ProvisioningTemplateApplyingInformation();
                ptai.HandlersToProcess ^=
                    OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.TermGroups;

                // Configure template parameters
                foreach (var key in template.Parameters.Keys)
                {
                    if (job.TemplateParameters.ContainsKey(key))
                    {
                        template.Parameters[key] = job.TemplateParameters[key];
                    }
                }

                // Apply the template to the target site
                web.ApplyProvisioningTemplate(template, ptai);

                Console.WriteLine("Applyed Provisioning Template \"{0}\" to site.",
                    job.ProvisioningTemplateUrl);
            }
        }
    }
}
