using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class ApplyProvisioningTemplateViewModel : JobViewModel
    {
        [Required(ErrorMessage = "Relative URL is a required field!")]
        [DisplayName("Relative URL")]
        public String RelativeUrl { get; set; }

        [Required(ErrorMessage = "Template is a required field!")]
        [DisplayName("Template")]
        [UIHint("ProvisioningTemplate")]
        public String ProvisioningTemplateUrl { get; set; }

        [DisplayName("Template Parameters")]
        public Dictionary<String, String> TemplateParameters { get; set; }
    }
}