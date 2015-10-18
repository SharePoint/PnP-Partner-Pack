using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    public static class PnPPartnerPackUtilities
    {
        public static Boolean IsPartnerPackOverridesEnabledOnSite(String siteUrl)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                String enabledString = context.Site.RootWeb.GetPropertyBagValueString(
                    PnPPartnerPackConstants.PnPPartnerPackOverridesPropertyBag, "false");
                return (Boolean.Parse(enabledString));
            }
        }

        public static void EnablePartnerPackOnSite(String siteUrl)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                ApplyProvisioningTemplateToSite(context,
                    HttpContext.Current.Server.MapPath("/Templates/Overrides"),
                    "PnP-Partner-Pack-Overrides.xml");

                // Turn ON the customization flag
                context.Site.RootWeb.SetPropertyBagValue(
                    PnPPartnerPackConstants.PnPPartnerPackOverridesPropertyBag, "true");
            }
        }

        public static void DisablePartnerPackOnSite(String siteUrl)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                // Remove the JavaScript injection overrides
                if (context.Site.ExistsJsLink(PnPPartnerPackConstants.PnPInjectedScriptName))
                {
                    context.Site.DeleteJsLink(PnPPartnerPackConstants.PnPInjectedScriptName);
                }

                // Turn OFF the customization flag
                context.Site.RootWeb.SetPropertyBagValue(
                    PnPPartnerPackConstants.PnPPartnerPackOverridesPropertyBag, "false");
            }
        }

        public static SiteCollectionSettings GetSiteCollectionSettings(String siteCollectionUri)
        {
            SiteCollectionSettings result = new SiteCollectionSettings();

            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteCollectionUri))
            {
                Web web = context.Web;
                context.Load(web, w => w.Title, w => w.Url);
                context.ExecuteQuery();

                result.Title = web.Title;
                result.Url = web.Url;
                result.PnPPartnerPackEnabled = PnPPartnerPackUtilities.IsPartnerPackOverridesEnabledOnSite(siteCollectionUri);
            }

            return (result);
        }

        private static void ApplyProvisioningTemplateToSite(ClientContext context, String container, String fileName)
        {
            // Configure the XML file system provider
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    container, "");

            // Load the template from the XML stored copy
            ProvisioningTemplate template = provider.GetTemplate(fileName);
            template.Connector = provider.Connector;

            // Apply the template to the target site
            context.Site.RootWeb.ApplyProvisioningTemplate(template);
        }

        public static void EnablePartnerPackInfrastructureOnSite(String siteUrl)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                try
                {
                    // Try to reference the target library
                    List templatesLibrary = context.Web.Lists.GetByTitle(
                        PnPPartnerPackConstants.PnPProvisioningTemplates);
                    context.Load(templatesLibrary);
                    context.ExecuteQuery();
                }
                catch (ServerException)
                {
                    ApplyProvisioningTemplateToSite(context,
                        HttpContext.Current.Server.MapPath("/Templates/Infrastructure"),
                        "PnP-Partner-Pack-Infrastructure-Templates.xml");
                }
            }
        }
    }
}