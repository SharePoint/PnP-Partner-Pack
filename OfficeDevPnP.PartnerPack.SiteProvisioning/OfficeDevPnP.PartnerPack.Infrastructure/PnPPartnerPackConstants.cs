using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    public static class PnPPartnerPackConstants
    {
        public static String ContentTypeIdField = "ContentTypeId";
        public static String TitleField = "Title";

        public static String PnPInjectedScriptName = "PnPPartnerPackOverrides";

        public static String PnPProvisioningTemplates = "PnPProvisioningTemplates";
        public static String PnPProvisioningTemplateContentTypeId = "0x010100498154DED7C84F4AB61A4029018F9048";
        public static String PnPProvisioningTemplateScope = "PnPProvisioningTemplateScope";
        public static String PnPProvisioningTemplatePlatform = "PnPProvisioningTemplatePlatform";
        public static String PnPProvisioningTemplateSourceUrl = "PnPProvisioningTemplateSourceUrl";

        public static String PnPProvisioningJobs = "PnPProvisioningJobs";
        public static String PnPProvisioningJobContentTypeId = "0x010100536B921A19A92949A056A9E7BEF008E5";
        public static String PnPProvisioningJobStatus = "PnPProvisioningJobStatus";
        public static String PnPProvisioningJobError = "PnPProvisioningJobError";
        public static String PnPProvisioningJobType = "PnPProvisioningJobType";
        public static String PnPProvisioningJobOwner = "PnPProvisioningJobOwner";

        public static String PnPPartnerPackOverridesPropertyBag = "_PnP_PartnerPack_Overrides_Enabled";

        public static String RegExSiteCollectionWithSubWebs = @"^(?<siteCollectionUrl>https\:\/\/(?<tenant>(\w|\-)+).sharepoint.com\/(sites|teams)\/(?<siteCollection>(\w|\-)+))(\/(?<subSite>(\w|\-)+))*";

        public const String TEMPLATE_SCOPE = "PnP_Template_Scope";
        public const String TEMPLATE_SCOPE_SITE = "Site";
        public const String TEMPLATE_SCOPE_WEB = "Web";
        public const String TEMPLATE_SCOPE_PARTIAL = "Partial";

        public const String PLATFORM_SPO = "PnP_Supports_SPO_Platform";
        public const String PLATFORM_SP2016 = "PnP_Supports_SP2016_Platform";
        public const String PLATFORM_SP2013 = "PnP_Supports_SP2013_Platform";

        public const String TRUE_VALUE = "True";

        public const String PropertyBag_Branding = "_PnP_Branding_TenantSettings";
        public const String PropertyBag_Branding_AppliedOn = "_PnP_Branding_AppliedOn";
        public const String PropertyBag_TemplateInfo = "_PnP_Template_Information";

        public const String BRANDING_SCRIPT_LINK_KEY = "PnP_Branding_ScriptLink";
    }
}