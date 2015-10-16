using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Components
{
    public static class PnPPartnerPackUtilities
    {
        public static Boolean IsPartnerPackEnabledOnSite(String siteUrl)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                String enabledString = context.Site.RootWeb.GetPropertyBagValueString("_PnP_PartnerPack_Enabled", "false");
                return (Boolean.Parse(enabledString));
            }
        }

        public static void EnablePartnerPackOnSite(String siteUrl)
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                ProvisioningTemplate template = new ProvisioningTemplate();
                template.CustomActions.SiteCustomActions.Add(
                    new CustomAction{
                        Name = PnPPartnerPackConstants.PnPInjectedScriptName,
                        Description = PnPPartnerPackConstants.PnPInjectedScriptName,
                        Location = "ScriptLink",
                        Sequence = 0,
                        ScriptSrc = PnPPartnerPackSettings.OverridesScriptUrl,
                        });

                context.Site.RootWeb.ApplyProvisioningTemplate(template);

                // Turn ON the customization flag
                context.Site.RootWeb.SetPropertyBagValue("_PnP_PartnerPack_Enabled", "false");
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
                context.Site.RootWeb.SetPropertyBagValue("_PnP_PartnerPack_Enabled", "false");
            }
        }
    }
}