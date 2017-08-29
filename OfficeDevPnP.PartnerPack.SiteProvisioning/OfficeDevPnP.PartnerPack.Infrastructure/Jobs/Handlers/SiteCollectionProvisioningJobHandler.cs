using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
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
                    using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContext())
                    {
                        adminContext.RequestTimeout = Timeout.Infinite;

                        // Configure the Site Collection properties
                        SiteEntity newSite = new SiteEntity();
                        newSite.Description = job.Description;
                        newSite.Lcid = (uint) job.Language;
                        newSite.Title = job.SiteTitle;
                        newSite.Url = siteUrl;
                        newSite.SiteOwnerLogin = job.PrimarySiteCollectionAdmin;
                        newSite.StorageMaximumLevel = job.StorageMaximumLevel;
                        newSite.StorageWarningLevel = job.StorageWarningLevel;

                        // Use the BaseSiteTemplate of the template, if any, otherwise 
                        // fallback to the pre-configured site template (i.e. STS#0)
                        newSite.Template = !String.IsNullOrEmpty(template.BaseSiteTemplate)
                            ? template.BaseSiteTemplate
                            : PnPPartnerPackSettings.DefaultSiteTemplate;

                        newSite.TimeZoneId = job.TimeZone;
                        newSite.UserCodeMaximumLevel = job.UserCodeMaximumLevel;
                        newSite.UserCodeWarningLevel = job.UserCodeWarningLevel;

                        // Create the Site Collection and wait for its creation (we're asynchronous)
                        var tenant = new Tenant(adminContext);
                        tenant.CreateSiteCollection(newSite, true, true); // TODO: Do we want to empty Recycle Bin?
                    }

                    // Work-around for the error 'Operation is not valid due to the current state of the object' (since DEC 2016)
                    // Get fresh Tenant instance after site collection was initially created.
                    using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContext())
                    {
                        var tenant = new Tenant(adminContext);
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
                        clientContext.RequestTimeout = Timeout.Infinite;

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

                        // We do intentionally remove taxonomies, which are not supported 
                        // in the AppOnly Authorization model
                        // For further details, see the PnP Partner Pack documentation 
                        ProvisioningTemplateApplyingInformation ptai =
                            new ProvisioningTemplateApplyingInformation();

                        // Write provisioning steps on console log
                        ptai.MessagesDelegate = (message, type) =>
                        {
                            switch (type)
                            {
                                case ProvisioningMessageType.Warning:
                                    {
                                        Console.WriteLine("{0} - {1}", type, message);
                                        break;
                                    }
                                case ProvisioningMessageType.Progress:
                                    {
                                        var activity = message;
                                        if (message.IndexOf("|") > -1)
                                        {
                                            var messageSplitted = message.Split('|');
                                            if (messageSplitted.Length == 4)
                                            {
                                                var status = messageSplitted[0];
                                                var statusDescription = messageSplitted[1];
                                                var current = double.Parse(messageSplitted[2]);
                                                var total = double.Parse(messageSplitted[3]);
                                                var percentage = Convert.ToInt32((100 / total) * current);
                                                Console.WriteLine("{0} - {1} - {2}", percentage, status, statusDescription);
                                            }
                                            else
                                            {
                                                Console.WriteLine(activity);
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine(activity);
                                        }
                                        break;
                                    }
                                case ProvisioningMessageType.Completed:
                                    {
                                        Console.WriteLine(type);
                                        break;
                                    }
                            }
                        };
                        ptai.ProgressDelegate = (message, step, total) =>
                        {
                            var percentage = Convert.ToInt32((100 / Convert.ToDouble(total)) * Convert.ToDouble(step));
                            Console.WriteLine("{0:00}/{1:00} - {2} - {3}", step, total, percentage, message);
                        };

                        // Exclude handlers not supported in App-Only
                        ptai.HandlersToProcess ^=
                            OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.TermGroups;
                        ptai.HandlersToProcess ^=
                            OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.SearchSettings;

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

                        // Fixup Title and Description
                        if (template.WebSettings != null)
                        {
                            template.WebSettings.Title = job.SiteTitle;
                            template.WebSettings.Description = job.Description;
                        }

                        // Replace existing Structural Current Navigation on target site
                        if (template.Navigation != null &&
                            template.Navigation.CurrentNavigation != null &&
                            template.Navigation.CurrentNavigation.StructuralNavigation != null &&
                            (template.Navigation.CurrentNavigation.NavigationType == CurrentNavigationType.Structural ||
                            template.Navigation.CurrentNavigation.NavigationType == CurrentNavigationType.StructuralLocal))
                        {
                            template.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes = true;
                        }
                        else if (template.Navigation != null &&
                            template.Navigation.CurrentNavigation != null &&
                            template.Navigation.CurrentNavigation.ManagedNavigation != null &&
                            template.Navigation.CurrentNavigation.NavigationType == CurrentNavigationType.Managed)
                        {
                            // We intentionally skip the Managed Navigation
                            template.Navigation = new Core.Framework.Provisioning.Model.Navigation(template.Navigation.GlobalNavigation, null);
                        }

                        // Replace existing Structural Global Navigation on target site
                        if (template.Navigation != null &&
                            template.Navigation.GlobalNavigation != null &&
                            template.Navigation.GlobalNavigation.StructuralNavigation != null &&
                            template.Navigation.GlobalNavigation.NavigationType == GlobalNavigationType.Structural)
                        {
                            template.Navigation.GlobalNavigation.StructuralNavigation.RemoveExistingNodes = true;
                        }
                        else if (template.Navigation != null &&
                            template.Navigation.GlobalNavigation != null &&
                            template.Navigation.GlobalNavigation.ManagedNavigation != null &&
                            template.Navigation.GlobalNavigation.NavigationType == GlobalNavigationType.Managed)
                        {
                            // We intentionally skip the Managed Navigation
                            template.Navigation = new Core.Framework.Provisioning.Model.Navigation(null, template.Navigation.CurrentNavigation);
                        }

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
                                if (brandingTemplate != null)
                                {
                                    if (brandingTemplate.WebSettings != null)
                                    {
                                        brandingTemplate.WebSettings.Title = job.SiteTitle;
                                        brandingTemplate.WebSettings.Description = job.Description;
                                    }

                                    // TO-DO: Need to handle exception here as there are multiple webs inside this where
                                    BrandingJobHandler.ApplyBrandingOnWeb(web, brandingSettings, brandingTemplate);
                                }
                            }
                        }


                        Console.WriteLine("Applied Provisioning Template \"{0}\" to site.",
                            job.ProvisioningTemplateUrl);
                    }
                }
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
