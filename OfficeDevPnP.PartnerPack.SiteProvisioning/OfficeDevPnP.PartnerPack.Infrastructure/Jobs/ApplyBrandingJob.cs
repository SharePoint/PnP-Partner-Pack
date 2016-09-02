using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines a Branding application Job
    /// </summary>
    public class ApplyBrandingJob : ProvisioningJob
    {
        /// <summary>
        ///  The Logo Image URL of the new branding, optional
        /// </summary>
        public String LogoImageUrl { get; set; }

        /// <summary>
        ///  The Background Image URL of the new branding, optional
        /// </summary>
        public String BackgroundImageUrl { get; set; }

        /// <summary>
        ///  The CSS Override URL of the new branding, optional
        /// </summary>
        public String CSSOverrideUrl { get; set; }

        /// <summary>
        ///  The Color File URL of the new branding, optional
        /// </summary>
        public String ColorFileUrl { get; set; }

        /// <summary>
        ///  The Palette File URL of the new branding, optional
        /// </summary>
        public String FontFileUrl { get; set; }

        /// <summary>
        ///  The UI Custom Actions URL of the new branding, optional
        /// </summary>
        public String UICustomActionsUrl { get; set; }
    }
}
