using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class BrandingViewModel : JobViewModel
    {
        [DisplayName("Logo Image File URL")]
        public String LogoImageUrl { get; set; }

        [DisplayName("Background Image File URL")]
        public String BackgroundImageUrl { get; set; }

        [DisplayName("CSS Overrides File URL")]
        public String CSSOverrideUrl { get; set; }

        [DisplayName("Color File URL")]
        public String ColorFileUrl { get; set; }

        [DisplayName("Font File URL")]
        public String FontFileUrl { get; set; }

        [DisplayName("UI Custom Actions File URL")]
        public String UICustomActionsUrl { get; set; }

        [DisplayName("Last Date and Time of Update")]
        public DateTime? UpdatedOn { get; set; }

        public Boolean RollOut { get; set; } = false;
    }
}