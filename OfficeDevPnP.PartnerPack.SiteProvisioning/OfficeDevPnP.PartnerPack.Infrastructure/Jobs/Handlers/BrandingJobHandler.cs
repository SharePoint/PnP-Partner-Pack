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
            ApplyBrandingJob brandingJob = job as ApplyBrandingJob;
            if (brandingJob == null)
            {
                throw new ArgumentException("Invalid job type for BrandingJobHandler.");
            }

            ApplyBranding(brandingJob);
        }

        private void ApplyBranding(ApplyBrandingJob job)
        {
            // Use the infrastructural site collection to temporary store the template with color palette and font files
            using (var repositoryContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(
                    PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                var repositoryWeb = repositoryContext.Site.RootWeb;
                repositoryContext.Load(repositoryWeb, w => w.Url);
                repositoryContext.ExecuteQueryRetry();

                var refererUri = new Uri(repositoryWeb.Url);
                var refererValue = $"{refererUri.Scheme}://{refererUri.Host}/";

                // Prepare an OpenXML provider
                XMLTemplateProvider provider = new XMLOpenXMLTemplateProvider($"{job.JobId}.pnp",
                    new SharePointConnector(repositoryContext, repositoryWeb.Url,
                            PnPPartnerPackConstants.PnPProvisioningTemplates));

                // Prepare the branding provisioning template
                var template = new ProvisioningTemplate();
                template.WebSettings = new WebSettings
                {
                    AlternateCSS = job.CSSOverrideUrl,
                    SiteLogo = job.LogoImageUrl,
                };

                template.ComposedLook = new ComposedLook();

                if (!String.IsNullOrEmpty(job.BackgroundImageUrl))
                {
                    var backgroundImageFileName = job.BackgroundImageUrl.Substring(job.BackgroundImageUrl.LastIndexOf("/") + 1);
                    var backgroundImageFileStream = HttpHelper.MakeGetRequestForStream(job.BackgroundImageUrl, "application/octet-stream", referer: refererValue);
                    template.ComposedLook.BackgroundFile = String.Format("{{sitecollection}}/SiteAssets/{0}", backgroundImageFileName);
                    provider.Connector.SaveFileStream(backgroundImageFileName, backgroundImageFileStream);

                    template.Files.Add(new Core.Framework.Provisioning.Model.File
                    {
                        Src = backgroundImageFileName,
                        Folder = "SiteAssets",
                        Overwrite = true,
                    });
                }

                if (!String.IsNullOrEmpty(job.FontFileUrl))
                {
                    var fontFileName = job.FontFileUrl.Substring(job.FontFileUrl.LastIndexOf("/") + 1);
                    var fontFileStream = HttpHelper.MakeGetRequestForStream(job.FontFileUrl, "application/octet-stream", referer : refererValue);
                    template.ComposedLook.FontFile = String.Format("{{themecatalog}}/15/{0}", fontFileName);
                    provider.Connector.SaveFileStream(fontFileName, fontFileStream);

                    template.Files.Add(new Core.Framework.Provisioning.Model.File {
                        Src = fontFileName,
                        Folder = "{themecatalog}/15",
                        Overwrite = true,
                    });
                }

                if (!String.IsNullOrEmpty(job.ColorFileUrl))
                {
                    var colorFileName = job.ColorFileUrl.Substring(job.ColorFileUrl.LastIndexOf("/") + 1);
                    var colorFileStream = HttpHelper.MakeGetRequestForStream(job.ColorFileUrl, "application/octet-stream", referer: refererValue);
                    template.ComposedLook.ColorFile = String.Format("{{themecatalog}}/15/{0}", colorFileName);
                    provider.Connector.SaveFileStream(colorFileName, colorFileStream);

                    template.Files.Add(new Core.Framework.Provisioning.Model.File
                    {
                        Src = colorFileName,
                        Folder = "{themecatalog}/15",
                        Overwrite = true,
                    });
                }

                // Save the template, ready to be applied
                provider.SaveAs(template, $"{job.JobId}.xml");

                // Re-open the template provider just saved
                provider = new XMLOpenXMLTemplateProvider($"{job.JobId}.pnp",
                    new SharePointConnector(repositoryContext, repositoryWeb.Url,
                            PnPPartnerPackConstants.PnPProvisioningTemplates));

                // Set the connector of the template, in order to being able to retrieve support files
                template.Connector = provider.Connector;

                // For each Site Collection in the tenant
                using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContext())
                {
                    var tenant = new Tenant(adminContext);

                    var siteCollections = tenant.GetSiteProperties(0, true);
                    adminContext.Load(siteCollections);
                    adminContext.ExecuteQueryRetry();

                    foreach (var site in siteCollections)
                    {
                        if (!site.Url.ToLower().Contains("/portals/"))
                        {
                            Console.WriteLine($"Applying branding to site: {site.Url}");

                            // Clean-up the template
                            template.WebSettings.MasterPageUrl = null;

                            using (var siteContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(site.Url))
                            {
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

                                // Apply branding (if it is not already applied)
                                var targetWeb = siteContext.Web;
                                targetWeb.EnsureProperty(w => w.MasterUrl);

                                // Check if we really need to apply/update the branding
                                var siteBrandingUpdatedOn = targetWeb.GetPropertyBagValueString(
                                    PnPPartnerPackConstants.PropertyBag_Branding_AppliedOn, null);

                                // If the branding updated on date and time are missing
                                // or older than the branding update date and time
                                if (String.IsNullOrEmpty(siteBrandingUpdatedOn) || 
                                    DateTime.Parse(siteBrandingUpdatedOn) < job.UpdatedOn.ToUniversalTime())
                                {
                                    // Confirm the master page URL
                                    template.WebSettings.MasterPageUrl = targetWeb.MasterUrl;

                                    // Apply the template
                                    targetWeb.ApplyProvisioningTemplate(template, ptai);

                                    // Apply a custom JSLink, if any
                                    if (!String.IsNullOrEmpty(job.UICustomActionsUrl))
                                    {
                                        targetWeb.AddJsLink(
                                            PnPPartnerPackConstants.BRANDING_SCRIPT_LINK_KEY,
                                            job.UICustomActionsUrl);
                                    }

                                    // Store Property Bag to set the last date and time when we applied the branding 
                                    targetWeb.SetPropertyBagValue(
                                        PnPPartnerPackConstants.PropertyBag_Branding_AppliedOn,
                                        DateTime.Now.ToUniversalTime().ToString());
                                }
                            }

                            Console.WriteLine($"Applied branding to site: {site.Url}");
                        }
                    }
                }

                provider.Delete($"{job.JobId}.pnp");
            }
        }
    }
}
