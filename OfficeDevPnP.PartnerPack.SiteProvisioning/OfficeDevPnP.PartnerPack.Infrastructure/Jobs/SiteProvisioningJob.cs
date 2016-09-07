using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Base abstract class for the Site Collection and Sub Site Provisioning Jobs
    /// </summary>
    public abstract class SiteProvisioningJob : ProvisioningJob
    {
        /// <summary>
        /// Defines the Title of the Site Collection or Sub Site to provision
        /// </summary>
        public String SiteTitle { get; set; }

        /// <summary>
        /// Defines the Relative URL of the Site Collection or Sub Site to provision
        /// </summary>
        public String RelativeUrl { get; set; }

        /// <summary>
        /// Declares the Site Policy for the Site Collection or Sub site to provision
        /// </summary>
        public String SitePolicy { get; set; }

        /// <summary>
        /// Defines the Description for the Site Collection or Sub site to provision
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Defines the Language for the Site Collection or Sub site to provision
        /// </summary>
        public Int32 Language { get; set; }

        /// <summary>
        /// Defines the Provisioning Template to use for the Site Collection or Sub site to provision
        /// </summary>
        public String ProvisioningTemplateUrl { get; set; }

        /// <summary>
        /// Defines the type name for the Templates Provider to use to get the selected template
        /// </summary>
        public String TemplatesProviderTypeName { get; set; }

        /// <summary>
        /// Defines the TimeZone for the Site Collection or Sub site to provision
        /// </summary>
        public Int32 TimeZone { get; set; }

        /// <summary>
        /// Defines the Parameters keys and values for the Site Collection or Sub site to provision
        /// </summary>
        public Dictionary<String, String> TemplateParameters { get; set; }

        /// <summary>
        /// Declares whether to apply tenant-level branding while creating the site
        /// </summary>
        public bool ApplyTenantBranding { get; set; } = false;
    }
}
