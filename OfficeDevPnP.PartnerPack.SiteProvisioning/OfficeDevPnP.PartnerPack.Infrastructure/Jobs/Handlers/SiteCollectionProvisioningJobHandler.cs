using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
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
    public class SiteCollectionProvisioningJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            SiteCollectionProvisioningJob scj = job as SiteCollectionProvisioningJob;
            if (scj == null)
            {
                throw new ArgumentException("Invalid job type for SiteCollectionProvisioningJobHandler.");
            }

            CreateSiteCollection(scj);
        }
        private void CreateSiteCollection(SiteCollectionProvisioningJob job)
        {
            Console.WriteLine("Creating Site Collection \"{0}\".", job.RelativeUrl);

            // Define the full Site Collection URL
            String siteUrl = String.Format("{0}{1}",
                PnPPartnerPackSettings.InfrastructureSiteUrl.Substring(0, PnPPartnerPackSettings.InfrastructureSiteUrl.IndexOf("sharepoint.com/") + 14),
                job.RelativeUrl);

            using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContext())
            {
                // Configure the Site Collection properties
                SiteEntity newSite = new SiteEntity();
                newSite.Description = job.Description;
                newSite.Lcid = (uint)job.Language;
                newSite.Title = job.SiteTitle;
                newSite.Url = siteUrl;
                newSite.SiteOwnerLogin = job.PrimarySiteCollectionAdmin;
                newSite.StorageMaximumLevel = job.StorageMaximumLevel;
                newSite.StorageWarningLevel = job.StorageWarningLevel;
                newSite.Template = PnPPartnerPackSettings.DefaultSiteTemplate;
                newSite.TimeZoneId = job.TimeZone;
                newSite.UserCodeMaximumLevel = job.UserCodeMaximumLevel;
                newSite.UserCodeWarningLevel = job.UserCodeWarningLevel;

                // Create the Site Collection and wait for its creation (we're asynchronous)
                var tenant = new Tenant(adminContext);
                tenant.CreateSiteCollection(newSite, true, true); // TODO: Do we want to empty Recycle Bin?

                Site site = tenant.GetSiteByUrl(siteUrl);
                Web web = site.RootWeb;

                adminContext.Load(site, s => s.Url);
                adminContext.Load(web, w => w.Url);
                adminContext.ExecuteQueryRetry();

                // Enable Secondary Site Collection Administrator
                if (!String.IsNullOrEmpty(job.SecondarySiteCollectionAdmin))
                {
                    Microsoft.SharePoint.Client.User secondaryOwner = web.EnsureUser(job.SecondarySiteCollectionAdmin);
                    secondaryOwner.IsSiteAdmin = true;
                    secondaryOwner.Update();

                    web.SiteUsers.AddUser(secondaryOwner);
                    adminContext.ExecuteQueryRetry();
                }

                Console.WriteLine("Site \"{0}\" created.", site.Url);

                // Check if external sharing has to be enabled
                if (job.ExternalSharingEnabled)
                {
                    EnableExternalSharing(tenant, site);

                    // Enable External Sharing
                    Console.WriteLine("Enabled External Sharing for site \"{0}\".",
                        site.Url);
                }
            }

            // Move to the context of the created Site Collection
            using (ClientContext clientContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                Site site = clientContext.Site;
                Web web = site.RootWeb;

                clientContext.Load(site, s => s.Url);
                clientContext.Load(web, w => w.Url);
                clientContext.ExecuteQueryRetry();

                // Check if we need to enable PnP Partner Pack overrides
                if (job.PartnerPackExtensionsEnabled)
                {
                    // Enable Responsive Design
                    PnPPartnerPackUtilities.EnablePartnerPackOnSite(site.Url);

                    Console.WriteLine("Enabled PnP Partner Pack Overrides on site \"{0}\".",
                        site.Url);
                }

                // Check if the site has to be responsive
                if (job.ResponsiveDesignEnabled)
                {
                    // Enable Responsive Design
                    PnPPartnerPackUtilities.EnableResponsiveDesignOnSite(site.Url);

                    Console.WriteLine("Enabled Responsive Design Template to site \"{0}\".",
                        site.Url);
                }

                // Apply the Provisioning Template
                Console.WriteLine("Applying Provisioning Template \"{0}\" to site.",
                    job.ProvisioningTemplateUrl);

                // Determine the reference URLs and file names
                String templatesSiteUrl = PnPPartnerPackUtilities.GetSiteCollectionRootUrl(job.ProvisioningTemplateUrl);
                String templateFileName = job.ProvisioningTemplateUrl.Substring(job.ProvisioningTemplateUrl.LastIndexOf("/") + 1);

                using (ClientContext repositoryContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(templatesSiteUrl))
                {
                    // Configure the XML file system provider
                    XMLTemplateProvider provider =
                        new XMLSharePointTemplateProvider(
                            repositoryContext,
                            templatesSiteUrl,
                            PnPPartnerPackConstants.PnPProvisioningTemplates);

                    // Load the template from the XML stored copy
                    ProvisioningTemplate template = provider.GetTemplate(templateFileName);
                    template.Connector = provider.Connector;

                    // We do intentionally remove taxonomies, which are not supported 
                    // in the AppOnly Authorization model
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

                    ptai.HandlersToProcess ^=
                        OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.TermGroups;

                    // Configure template parameters
                    if (job.TemplateParameters != null)
                    {
                        foreach (var key in job.TemplateParameters.Keys)
                        {
                            if (job.TemplateParameters.ContainsKey(key))
                            {
                                template.Parameters[key] = job.TemplateParameters[key];
                            }
                        }
                    }

                    web.ApplyProvisioningTemplate(template, ptai);
                }

                Console.WriteLine("Applyed Provisioning Template \"{0}\" to site.",
                    job.ProvisioningTemplateUrl);
            }
        }

        public void EnableExternalSharing(Tenant tenant, Site site)
        {
            ClientContext context = (ClientContext)site.Context;
            tenant.SetSiteProperties(site.Url, sharingCapability: SharingCapabilities.ExternalUserSharingOnly);

            SiteProperties siteProps = tenant.GetSitePropertiesByUrl(site.Url, false);
            context.Load(tenant);
            context.Load(siteProps);
            context.ExecuteQuery();
        }
    }
}
