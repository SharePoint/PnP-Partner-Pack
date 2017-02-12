using OfficeDevPnP.PartnerPack.Infrastructure;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    /// <summary>
    /// Basic abstract class for a Site Provisioning View Model
    /// </summary>
    public abstract class CreateSiteViewModel : JobViewModel
    {
        public CreateSiteStep Step { get; set; }

        [Required(ErrorMessageResourceName = "TitleRequired", ErrorMessageResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [Display(Name = "Title", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        public String Title { get; set; }

        [Required(ErrorMessageResourceName = "RelativeURLRequired", ErrorMessageResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [MaxLength(33)]
        [Display(Name = "RelativeURL", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        public String RelativeUrl { get; set; }

        [Display(Name = "SitePolicy", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [UIHint("SitePolicy")]
        public String SitePolicy { get; set; }

        [Display(Name = "SiteDescription", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [DataType(DataType.MultilineText)]
        [UIHint("Multilines")]
        public string Description { get; set; }

        [Required(ErrorMessageResourceName = "SiteLanguageRequired", ErrorMessageResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [Display(Name = "SiteLanguage", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [UIHint("LocaleID")]
        public Int32 Language { get; set; }

        [Required(ErrorMessageResourceName = "SiteTemplateRequired", ErrorMessageResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [MinLength(1)]
        [Display(Name = "SiteTemplate", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [UIHint("ProvisioningTemplate")]
        public String ProvisioningTemplateUrl { get; set; }

        [Display(Name = "TimeZone", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        [UIHint("TimeZoneInfo")]
        public Int32 TimeZone { get; set; }

        [Display(Name = "TemplateParameters", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        public Dictionary<String, String> TemplateParameters { get; set; }

        public TargetScope Scope { get; set; }

        public String TemplatesProviderTypeName { get; set; }

        public String ParentSiteUrl { get; set; }

        [Display(Name = "ApplyTenantBranding", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        public Boolean ApplyTenantBranding { get; set; }
    }

    public enum CreateSiteStep
    {
        TemplateSelection,
        SiteInformation,
        TemplateParameters,
        SiteCreated,
    }
}