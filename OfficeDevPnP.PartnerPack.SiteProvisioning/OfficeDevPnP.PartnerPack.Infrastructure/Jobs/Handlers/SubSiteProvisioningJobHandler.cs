using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
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
                context.RequestTimeout = Timeout.Infinite;

                // Get a reference to the parent Web
                Web parentWeb = context.Web;

                // Load the template from the source Templates Provider
                if (!String.IsNullOrEmpty(job.TemplatesProviderTypeName))
                {
                    ProvisioningTemplate template = null;

                    var templatesProvider = PnPPartnerPackSettings.TemplatesProviders[job.TemplatesProviderTypeName];
                    if (templatesProvider != null)
                    {
                        template = templatesProvider.GetProvisioningTemplate(job.ProvisioningTemplateUrl);
                    }

                    if (template != null)
                    {
                        // Create the new sub site as a new child Web
                        WebCreationInformation newWeb = new WebCreationInformation();
                        newWeb.Description = job.Description;
                        newWeb.Language = job.Language;
                        newWeb.Title = job.SiteTitle;
                        newWeb.Url = subSiteUrl;
                        newWeb.UseSamePermissionsAsParentSite = job.InheritPermissions;

                        // Use the BaseSiteTemplate of the template, if any, otherwise 
                        // fallback to the pre-configured site template (i.e. STS#0)
                        newWeb.WebTemplate = !String.IsNullOrEmpty(template.BaseSiteTemplate) ? 
                            template.BaseSiteTemplate : 
                            PnPPartnerPackSettings.DefaultSiteTemplate;

                        Web web = parentWeb.Webs.Add(newWeb);
                        context.ExecuteQueryRetry();

                        Console.WriteLine("Site \"{0}\" created.", subSiteUrl);

                        // Apply the Provisioning Template
                        Console.WriteLine("Applying Provisioning Template \"{0}\" to site.",
                            job.ProvisioningTemplateUrl);

                        // We do intentionally remove taxonomies, which are not supported in the AppOnly Authorization model
                        // For further details, see the PnP Partner Pack documentation 
                        ProvisioningTemplateApplyingInformation ptai =
                            new ProvisioningTemplateApplyingInformation();

                        // Write provisioning steps on console log
                        ptai.MessagesDelegate += delegate (string message, ProvisioningMessageType messageType) {
                            Console.WriteLine("{0} - {1}", messageType, messageType);
                        };
                        ptai.ProgressDelegate += delegate (string message, int step, int total) {
                            Console.WriteLine("{0:00}/{1:00} - {2}", step, total, message);
                        };

                        // Exclude handlers not supported in App-Only
                        ptai.HandlersToProcess ^=
                            OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.TermGroups;
                        ptai.HandlersToProcess ^=
                            OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.SearchSettings;

                        // Configure template parameters
                        foreach (var key in template.Parameters.Keys)
                        {
                            if (job.TemplateParameters.ContainsKey(key))
                            {
                                template.Parameters[key] = job.TemplateParameters[key];
                            }
                        }

                        // Fixup Title and Description
                        template.WebSettings.Title = job.SiteTitle;
                        template.WebSettings.Description = job.Description;

                        // Apply the template to the target site
                        web.ApplyProvisioningTemplate(template, ptai);

                        // Save the template information in the target site
                        var info = new SiteTemplateInfo()
                        {
                            TemplateProviderType = job.TemplatesProviderTypeName,
                            TemplateUri = job.ProvisioningTemplateUrl,
                            TemplateParameters = template.Parameters,
                            AppliedOn = DateTime.Now,
                        };
                        var jsonInfo = JsonConvert.SerializeObject(info);
                        web.SetPropertyBagValue(PnPPartnerPackConstants.PropertyBag_TemplateInfo, jsonInfo);

                        // Set site policy template
                        if (!String.IsNullOrEmpty(job.SitePolicy))
                        {
                            web.ApplySitePolicy(job.SitePolicy);
                        }

                        // Apply Tenant Branding, if requested
                        if (job.ApplyTenantBranding)
                        {
                            var brandingSettings = PnPPartnerPackUtilities.GetTenantBrandingSettings();

                            using (var repositoryContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(
                                PnPPartnerPackSettings.InfrastructureSiteUrl))
                            {
                                var brandingTemplate = BrandingJobHandler.PrepareBrandingTemplate(repositoryContext, brandingSettings);

                                // Fixup Title and Description
                                brandingTemplate.WebSettings.Title = job.SiteTitle;
                                brandingTemplate.WebSettings.Description = job.Description;

                                BrandingJobHandler.ApplyBrandingOnWeb(web, brandingSettings, brandingTemplate);
                            }
                        }

                        Console.WriteLine("Applyed Provisioning Template \"{0}\" to site.",
                            job.ProvisioningTemplateUrl);
                    }
                }
            }
        }
    }
}
