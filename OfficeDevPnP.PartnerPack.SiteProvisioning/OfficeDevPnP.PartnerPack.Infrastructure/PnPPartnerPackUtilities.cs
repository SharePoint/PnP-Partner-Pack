using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
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
                    PnPPartnerPackSettings.InfrastructureSiteUrl,
                    "Overrides",
                    "PnP-Partner-Pack-Overrides.xml",
                    handlers: Handlers.CustomActions);

                // Turn ON the customization flag
                context.Site.RootWeb.SetPropertyBagValue(
                    PnPPartnerPackConstants.PnPPartnerPackOverridesPropertyBag, "true");
            }
        }

        public static void EnableResponsiveDesignOnSite(String siteUrl)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                Web web = context.Web;
                Site site = context.Site;

                if (web.IsSubSite())
                {
                    web.EnableResponsiveUI(PnPPartnerPackSettings.InfrastructureSiteUrl);
                }
                else
                {
                    site.EnableResponsiveUI(PnPPartnerPackSettings.InfrastructureSiteUrl);
                }
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

        public static String GetPropertyBagValueFromInfrastructure(String propertyKey)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(
                PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                var web = context.Web;
                var result = web.GetPropertyBagValueString(propertyKey, null);

                return (result);
            }
        }

        public static void SetPropertyBagValueToInfrastructure(String propertyKey, String value)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(
                PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                var web = context.Web;
                web.SetPropertyBagValue(propertyKey, value);
            }
        }

        private static void ApplyProvisioningTemplateToSite(ClientContext context, String siteUrl, String folder, String fileName, Dictionary<String, String> parameters = null, Handlers handlers = Handlers.All)
        {
            // TODO: Move to Open XML

            // Configure the XML file system provider
            XMLTemplateProvider provider =
                new XMLSharePointTemplateProvider(context, siteUrl,
                    PnPPartnerPackConstants.PnPProvisioningTemplates +
                    (!String.IsNullOrEmpty(folder) ? "/" + folder : String.Empty));

            // Load the template from the XML stored copy
            ProvisioningTemplate template = provider.GetTemplate(fileName);
            template.Connector = provider.Connector;

            ProvisioningTemplateApplyingInformation ptai =
                new ProvisioningTemplateApplyingInformation();

            // We exclude Term Groups because they are not supported in AppOnly
            ptai.HandlersToProcess = handlers;
            ptai.HandlersToProcess ^= Handlers.TermGroups;

            // Handle any custom parameter
            if (parameters != null)
            {
                foreach (var parameter in parameters)
                {
                    template.Parameters.Add(parameter.Key, parameter.Value);
                }
            }

            // Apply the template to the target site
            context.Site.RootWeb.ApplyProvisioningTemplate(template, ptai);
        }

        public static void EnablePartnerPackInfrastructureOnSite(String siteUrl)
        {
            siteUrl = PnPPartnerPackUtilities.GetSiteCollectionRootUrl(siteUrl);

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
                        PnPPartnerPackSettings.InfrastructureSiteUrl,
                        "Infrastructure",
                        "PnP-Partner-Pack-Infrastructure-Templates.xml");
                }
            }
        }

        public static Boolean UserIsTenantGlobalAdmin()
        {
            return (UserIsAdmin(MicrosoftGraphConstants.GlobalTenantAdminRole));
        }

        public static Boolean UserIsSPOAdmin()
        {
            return (UserIsAdmin(MicrosoftGraphConstants.GlobalSPOAdminRole));
        }

        private static Boolean UserIsAdmin(String targetRole)
        {
            try
            {
                // Retrieve (using the Microsoft Graph) the current user's roles
                String jsonResponse = HttpHelper.MakeGetRequestForString(
                    String.Format("{0}me/memberOf?$select=id,displayName",
                        MicrosoftGraphConstants.MicrosoftGraphV1BaseUri),
                    MicrosoftGraphHelper.GetAccessTokenForCurrentUser(MicrosoftGraphConstants.MicrosoftGraphResourceId));

                if (jsonResponse != null)
                {
                    var result = JsonConvert.DeserializeObject<UserRoles>(jsonResponse);
                    // Check if the requested role (of type DirectoryRole) is included in the list
                    return (result.Roles.Any(r => r.DisplayName == targetRole && 
                        r.DataType.Equals("#microsoft.graph.directoryRole", StringComparison.InvariantCultureIgnoreCase)));
                }
            }
            catch (Exception)
            {
                // Ignore any exception and return false (user is not member of ...)
            }

            return (false);
        }

        public static String GetSiteCollectionRootUrl(String siteUrl)
        {
            // Get the Site Collection root URL
            System.Text.RegularExpressions.Regex regex =
                new System.Text.RegularExpressions.Regex(
                    PnPPartnerPackConstants.RegExSiteCollectionWithSubWebs);

            var match = regex.Match(siteUrl);
            if (match.Success)
            {
                siteUrl = match.Groups["siteCollectionUrl"].Value;
            }

            return (siteUrl);
        }

        /// <summary>
        /// Helper method for reading branding settings from the Infrastructural Site Collection
        /// </summary>
        /// <returns></returns>
        public static BrandingSettings GetTenantBrandingSettings()
        {
            // Get the current settings from the Infrastructural Site Collection
            var jsonBrandingSettings = PnPPartnerPackUtilities.GetPropertyBagValueFromInfrastructure(
                PnPPartnerPackConstants.PropertyBag_Branding);

            // Read the current branding settings, if any
            var branding = jsonBrandingSettings != null ?
                JsonConvert.DeserializeObject<BrandingSettings>(jsonBrandingSettings) :
                new BrandingSettings();

            return branding;
        }
    }
}