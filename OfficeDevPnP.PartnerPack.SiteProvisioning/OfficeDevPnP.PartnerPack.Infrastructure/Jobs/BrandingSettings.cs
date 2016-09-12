using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines the settings for Tenant branding
    /// </summary>
    public class BrandingSettings
    {
        /// <summary>
        ///  Represents the URL of the Logo, if any
        /// </summary>
        public String LogoImageUrl { get; set; }

        /// <summary>
        /// Represents the URL of the background image, if any
        /// </summary>
        public String BackgroundImageUrl { get; set; }

        /// <summary>
        /// Represents the URL of CSS override file, if any
        /// </summary>
        public String CSSOverrideUrl { get; set; }

        /// <summary>
        /// Represents the URL of the Color file, if any
        /// </summary>
        public String ColorFileUrl { get; set; }

        /// <summary>
        /// Represents the URL of the Font file, if any
        /// </summary>
        public String FontFileUrl { get; set; }

        /// <summary>
        /// Represents the URL of the UI Custom action, if any
        /// </summary>
        public String UICustomActionsUrl { get; set; }

        /// <summary>
        /// Latest date and time of update for the branding settings
        /// </summary>
        public DateTime? UpdatedOn { get; set; }
    }
}
