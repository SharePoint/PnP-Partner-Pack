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

        [Required(ErrorMessage = "Title is a required field!")]
        [DisplayName("Title")]
        public String Title { get; set; }

        [Required(ErrorMessage = "Relative URL is a required field!")]
        [DisplayName("Relative URL")]
        public String RelativeUrl { get; set; }

        [DisplayName("Site Policy")]
        [UIHint("SitePolicy")]
        public String SitePolicy { get; set; }

        [DisplayName("Description")]
        [DataType(DataType.MultilineText)]
        public string Description { get; set; }

        [Required(ErrorMessage = "Language is a required field!")]
        [DisplayName("Language")]
        [UIHint("LocaleID")]
        public Int32 Language { get; set; }

        [Required(ErrorMessage = "Template is a required field!")]
        [MinLength(1)]
        [DisplayName("Template")]
        [UIHint("ProvisioningTemplate")]
        public String ProvisioningTemplateUrl { get; set; }

        [DisplayName("Time Zone")]
        [UIHint("TimeZoneInfo")]
        public Int32 TimeZone { get; set; }

        [DisplayName("Template Parameters")]
        public Dictionary<String, String> TemplateParameters { get; set; }

        public TemplateScope Scope { get; set; }

        public String ParentSiteUrl { get; set; }
    }

    public enum CreateSiteStep
    {
        TemplateSelection,
        SiteInformation,
        TemplateParameters,
        SiteCreated,
    }
}