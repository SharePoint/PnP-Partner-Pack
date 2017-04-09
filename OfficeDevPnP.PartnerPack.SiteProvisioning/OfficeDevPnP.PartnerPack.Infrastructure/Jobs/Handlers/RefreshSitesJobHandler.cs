using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    public class RefreshSitesJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            RefreshSitesJob updateJob = job as RefreshSitesJob;
            if (updateJob == null)
            {
                throw new ArgumentException("Invalid job type for RefreshSitesJobHandler.");
            }

            UpdateTemplates(updateJob);
        }

        private void UpdateTemplates(RefreshSitesJob job)
        {
            // For each Site Collection in the tenant
            using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContext())
            {
                var tenant = new Tenant(adminContext);

                var siteCollections = tenant.GetSiteProperties(0, true);
                adminContext.Load(siteCollections);
                adminContext.ExecuteQueryRetry();

                foreach (var site in siteCollections)
                {
                    // Exclude Microsoft NextGen Portals
                    if (!site.Url.ToLower().Contains("/portals/") 
                        && !site.Url.ToLower().Contains("-public.sharepoint.com") 
                        && !site.Url.ToLower().Contains("-my.sharepoint.com"))
                    {
                        using (var siteContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(site.Url))
                        {
                            // Get a reference to the target web
                            var targetWeb = siteContext.Site.RootWeb;

                            // Update the root web of the site collection
                            RefreshSitesJobHandler.UpdateTemplateOnWeb(targetWeb, job);
                        }
                    }
                }
            }
        }

        internal static void UpdateTemplateOnWeb(Web targetWeb, RefreshSitesJob job = null)
        {
            targetWeb.EnsureProperty(w => w.Url);

            var infoJson = targetWeb.GetPropertyBagValueString(PnPPartnerPackConstants.PropertyBag_TemplateInfo, null);
            if (!String.IsNullOrEmpty(infoJson))
            {
                Console.WriteLine($"Updating template for site: {targetWeb.Url}");

                var info = JsonConvert.DeserializeObject<SiteTemplateInfo>(infoJson);

                // If we have the template info
                if (info != null && !String.IsNullOrEmpty(info.TemplateProviderType))
                {
                    ProvisioningTemplate template = null;

                    // Try to retrieve the template
                    var templatesProvider = PnPPartnerPackSettings.TemplatesProviders[info.TemplateProviderType];
                    if (templatesProvider != null)
                    {
                        template = templatesProvider.GetProvisioningTemplate(info.TemplateUri);
                    }

                    // If we have the template
                    if (template != null)
                    {
                        // Configure proper settings for the provisioning engine
                        ProvisioningTemplateApplyingInformation ptai =
                            new ProvisioningTemplateApplyingInformation();

                        // Write provisioning steps on console log
                        ptai.MessagesDelegate += delegate (string message, ProvisioningMessageType messageType)
                        {
                            Console.WriteLine("{0} - {1}", messageType, messageType);
                        };
                        ptai.ProgressDelegate += delegate (string message, int step, int total)
                        {
                            Console.WriteLine("{0:00}/{1:00} - {2}", step, total, message);
                        };

                        // Exclude handlers not supported in App-Only
                        ptai.HandlersToProcess ^=
                            OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.TermGroups;
                        ptai.HandlersToProcess ^=
                            OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.SearchSettings;

                        // Configure template parameters
                        if (info.TemplateParameters != null)
                        {
                            foreach (var key in info.TemplateParameters.Keys)
                            {
                                if (info.TemplateParameters.ContainsKey(key))
                                {
                                    template.Parameters[key] = info.TemplateParameters[key];
                                }
                            }
                        }

                        targetWeb.ApplyProvisioningTemplate(template, ptai);

                        // Save the template information in the target site
                        var updatedInfo = new SiteTemplateInfo()
                        {
                            TemplateProviderType = info.TemplateProviderType,
                            TemplateUri = info.TemplateUri,
                            TemplateParameters = template.Parameters,
                            AppliedOn = DateTime.Now,
                        };
                        var jsonInfo = JsonConvert.SerializeObject(updatedInfo);
                        targetWeb.SetPropertyBagValue(PnPPartnerPackConstants.PropertyBag_TemplateInfo, jsonInfo);

                        Console.WriteLine($"Updated template on site: {targetWeb.Url}");

                        // Update (recursively) all the subwebs of the current web
                        targetWeb.EnsureProperty(w => w.Webs);

                        foreach (var subweb in targetWeb.Webs)
                        {
                            UpdateTemplateOnWeb(subweb, job);
                        }
                    }
                }
            }
        }
    }
}
