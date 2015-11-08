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
        public static String PnPProvisioningTemplateSourceUrl = "PnPProvisioningTemplateSourceUrl";

        public static String PnPProvisioningJobs = "PnPProvisioningJobs";
        public static String PnPProvisioningJobContentTypeId = "0x010100536B921A19A92949A056A9E7BEF008E5";
        public static String PnPProvisioningJobStatus = "PnPProvisioningJobStatus";
        public static String PnPProvisioningJobError = "PnPProvisioningJobError";
        public static String PnPProvisioningJobType = "PnPProvisioningJobType";
        public static String PnPProvisioningJobOwner = "PnPProvisioningJobOwner";

        public static String PnPPartnerPackOverridesPropertyBag = "_PnP_PartnerPack_Overrides_Enabled";

        public static String RegExSiteCollectionWithSubWebs = @"^(?<siteCollectionUrl>https\:\/\/(?<tenant>(\w|\-)+).sharepoint.com\/(sites|teams)\/(?<siteCollection>(\w|\-)+))(\/(?<subSite>(\w|\-)+))*";
    }
}