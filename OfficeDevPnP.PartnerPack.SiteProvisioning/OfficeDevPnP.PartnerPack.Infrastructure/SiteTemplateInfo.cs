using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Defines the information about a template applied to a site
    /// </summary>
    public class SiteTemplateInfo
    {
        /// <summary>
        /// The template provider
        /// </summary>
        public String TemplateProviderType { get; set; }

        /// <summary>
        /// The URI of the applied template
        /// </summary>
        public String TemplateUri { get; set; }

        /// <summary>
        /// The latest date of application of the template
        /// </summary>
        public DateTime AppliedOn { get; set; }
    }
}
