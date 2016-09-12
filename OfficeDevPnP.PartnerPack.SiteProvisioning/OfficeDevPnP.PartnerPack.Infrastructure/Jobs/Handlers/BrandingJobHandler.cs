using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
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
    public class BrandingJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            BrandingJob brandingJob = job as BrandingJob;
            if (brandingJob == null)
            {
                throw new ArgumentException("Invalid job type for BrandingJobHandler.");
            }

            ApplyBranding(brandingJob);
        }

        private void ApplyBranding(BrandingJob job)
        {
            // Use the infrastructural site collection to temporary store the template with color palette and font files
            using (var repositoryContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(
                    PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                var brandingSettings = PnPPartnerPackUtilities.GetTenantBrandingSettings();
                ProvisioningTemplate template = PrepareBrandingTemplate(repositoryContext, brandingSettings);

                // For each Site Collection in the tenant
                using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContext())
                {
                    var tenant = new Tenant(adminContext);

                    var siteCollections = tenant.GetSiteProperties(0, true);
                    adminContext.Load(siteCollections);
                    adminContext.ExecuteQueryRetry();

                    foreach (var site in siteCollections)
                    {
                        if (!site.Url.ToLower().Contains("/portals/")
                            && !site.Url.ToLower().Contains("-public.sharepoint.com")
                            && !site.Url.ToLower().Contains("-my.sharepoint.com"))
                        {
                            // Clean-up the template
                            template.WebSettings.MasterPageUrl = null;

                            using (var siteContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(site.Url))
                            {
                                // Get a reference to the target web
                                var targetWeb = siteContext.Site.RootWeb;

                                // Update the root web of the site collection
                                ApplyBrandingOnWeb(targetWeb, brandingSettings, template);
                            }
                        }
                    }
                }
            }
        }

        public static ProvisioningTemplate PrepareBrandingTemplate(ClientContext repositoryContext, BrandingSettings brandingSettings)
        {
            var repositoryWeb = repositoryContext.Site.RootWeb;
            repositoryContext.Load(repositoryWeb, w => w.Url);
            repositoryContext.ExecuteQueryRetry();

            var refererUri = new Uri(repositoryWeb.Url);
            var refererValue = $"{refererUri.Scheme}://{refererUri.Host}/";

            var templateId = Guid.NewGuid();

            // Prepare an OpenXML provider
            XMLTemplateProvider provider = new XMLOpenXMLTemplateProvider($"{templateId}.pnp",
                new SharePointConnector(repositoryContext, repositoryWeb.Url,
                        PnPPartnerPackConstants.PnPProvisioningTemplates));

            // Prepare the branding provisioning template
            var template = new ProvisioningTemplate()
            {
                Id = $"Branding-{Guid.NewGuid()}",
                DisplayName = "Branding Template",
            };

            template.WebSettings = new WebSettings
            {
                AlternateCSS = brandingSettings.CSSOverrideUrl,
                SiteLogo = brandingSettings.LogoImageUrl,
            };

            template.ComposedLook = new ComposedLook()
            {
                Name = "SharePointBranding",
            };

            if (!String.IsNullOrEmpty(brandingSettings.BackgroundImageUrl))
            {
                var backgroundImageFileName = brandingSettings.BackgroundImageUrl.Substring(brandingSettings.BackgroundImageUrl.LastIndexOf("/") + 1);
                var backgroundImageFileStream = HttpHelper.MakeGetRequestForStream(brandingSettings.BackgroundImageUrl, "application/octet-stream", referer: refererValue);
                template.ComposedLook.BackgroundFile = String.Format("{{sitecollection}}/SiteAssets/{0}", backgroundImageFileName);
                provider.Connector.SaveFileStream(backgroundImageFileName, backgroundImageFileStream);

                template.Files.Add(new Core.Framework.Provisioning.Model.File
                {
                    Src = backgroundImageFileName,
                    Folder = "SiteAssets",
                    Overwrite = true,
                });
            }
            else
            {
                template.ComposedLook.BackgroundFile = String.Empty;
            }

            if (!String.IsNullOrEmpty(brandingSettings.FontFileUrl))
            {
                var fontFileName = brandingSettings.FontFileUrl.Substring(brandingSettings.FontFileUrl.LastIndexOf("/") + 1);
                var fontFileStream = HttpHelper.MakeGetRequestForStream(brandingSettings.FontFileUrl, "application/octet-stream", referer: refererValue);
                template.ComposedLook.FontFile = String.Format("{{themecatalog}}/15/{0}", fontFileName);
                provider.Connector.SaveFileStream(fontFileName, fontFileStream);

                template.Files.Add(new Core.Framework.Provisioning.Model.File
                {
                    Src = fontFileName,
                    Folder = "{themecatalog}/15",
                    Overwrite = true,
                });
            }
            else
            {
                template.ComposedLook.FontFile = String.Empty;
            }

            if (!String.IsNullOrEmpty(brandingSettings.ColorFileUrl))
            {
                var colorFileName = brandingSettings.ColorFileUrl.Substring(brandingSettings.ColorFileUrl.LastIndexOf("/") + 1);
                var colorFileStream = HttpHelper.MakeGetRequestForStream(brandingSettings.ColorFileUrl, "application/octet-stream", referer: refererValue);
                template.ComposedLook.ColorFile = String.Format("{{themecatalog}}/15/{0}", colorFileName);
                provider.Connector.SaveFileStream(colorFileName, colorFileStream);

                template.Files.Add(new Core.Framework.Provisioning.Model.File
                {
                    Src = colorFileName,
                    Folder = "{themecatalog}/15",
                    Overwrite = true,
                });
            }
            else
            {
                template.ComposedLook.ColorFile = String.Empty;
            }

            // Save the template, ready to be applied
            provider.SaveAs(template, $"{templateId}.xml");

            // Re-open the template provider just saved
            provider = new XMLOpenXMLTemplateProvider($"{templateId}.pnp",
                new SharePointConnector(repositoryContext, repositoryWeb.Url,
                        PnPPartnerPackConstants.PnPProvisioningTemplates));

            // Set the connector of the template, in order to being able to retrieve support files
            template.Connector = provider.Connector;

            return template;
        }

        public static void ApplyBrandingOnWeb(Web targetWeb, BrandingSettings brandingSettings, ProvisioningTemplate template)
        {
            targetWeb.EnsureProperties(w => w.MasterUrl, w => w.Url);

            // Configure proper settings for the provisioning engine
            ProvisioningTemplateApplyingInformation ptai =
                new ProvisioningTemplateApplyingInformation();

            // Write provisioning steps on console log
            ptai.MessagesDelegate += delegate (string message, ProvisioningMessageType messageType) {
                Console.WriteLine("{0} - {1}", messageType, messageType);
            };
            ptai.ProgressDelegate += delegate (string message, int step, int total) {
                Console.WriteLine("{0:00}/{1:00} - {2}", step, total, message);
            };

            // Include only required handlers
            ptai.HandlersToProcess = Core.Framework.Provisioning.Model.Handlers.ComposedLook |
                Core.Framework.Provisioning.Model.Handlers.Files |
                Core.Framework.Provisioning.Model.Handlers.WebSettings;

            // Check if we really need to apply/update the branding
            var siteBrandingUpdatedOn = targetWeb.GetPropertyBagValueString(
                PnPPartnerPackConstants.PropertyBag_Branding_AppliedOn, null);

            // If the branding updated on date and time are missing
            // or older than the branding update date and time
            if (String.IsNullOrEmpty(siteBrandingUpdatedOn) ||
                DateTime.Parse(siteBrandingUpdatedOn) < brandingSettings.UpdatedOn.Value.ToUniversalTime())
            {
                Console.WriteLine($"Appling branding to site: {targetWeb.Url}");

                // Confirm the master page URL
                template.WebSettings.MasterPageUrl = targetWeb.MasterUrl;

                // Apply the template
                targetWeb.ApplyProvisioningTemplate(template, ptai);

                // Apply a custom JSLink, if any
                if (!String.IsNullOrEmpty(brandingSettings.UICustomActionsUrl))
                {
                    targetWeb.AddJsLink(
                        PnPPartnerPackConstants.BRANDING_SCRIPT_LINK_KEY,
                        brandingSettings.UICustomActionsUrl);
                }

                // Store Property Bag to set the last date and time when we applied the branding 
                targetWeb.SetPropertyBagValue(
                    PnPPartnerPackConstants.PropertyBag_Branding_AppliedOn,
                    DateTime.Now.ToUniversalTime().ToString());

                Console.WriteLine($"Applied branding to site: {targetWeb.Url}");
            }

            // Apply branding (recursively) on all the subwebs of the current web
            targetWeb.EnsureProperty(w => w.Webs);

            foreach (var subweb in targetWeb.Webs)
            {
                ApplyBrandingOnWeb(subweb, brandingSettings, template);
            }
        }
    }
}
