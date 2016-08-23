using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines a Provisioning Template extraction Job
    /// </summary>
    public class GetProvisioningTemplateJob : ProvisioningJob
    {
        /// <summary>
        ///  The Description of the Provisioning Template that will be saved
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// The filename to use for storing the Provisioning Template
        /// </summary>
        public String FileName { get; set; }

        /// <summary>
        /// Defines whether to include all term groups during extraction of the Provisioning Template
        /// </summary>
        public Boolean IncludeAllTermGroups { get; set; }

        /// <summary>
        /// Defines whether to include site collection level term groups during extraction of the Provisioning Template
        /// </summary>
        public Boolean IncludeSiteCollectionTermGroup { get; set; }

        /// <summary>
        /// Defines whether to include search settings during extraction of the Provisioning Template
        /// </summary>
        public Boolean IncludeSearchConfiguration { get; set; }

        /// <summary>
        /// Defines whether to include site groups during extraction of the Provisioning Template
        /// </summary>
        public Boolean IncludeSiteGroups { get; set; }

        /// <summary>
        /// Defines whether to persist the composed look files during extraction of the Provisioning Template
        /// </summary>
        public Boolean PersistComposedLookFiles { get; set; }

        /// <summary>
        /// Defines the Source Site URL of the Provisioning Template
        /// </summary>
        public String SourceSiteUrl { get; set; }

        /// <summary>
        /// Defines the Scope of the Provisioning Template
        /// </summary>
        public TargetScope Scope { get; set; }

        /// <summary>
        /// Defines the target Location for the Provisioning Template file
        /// </summary>
        public ProvisioningTemplateLocation Location { get; set; }

        /// <summary>
        /// Defines the URL of the Site that will store the Provisioning Template file
        /// </summary>
        public String StorageSiteLocationUrl { get; set; }

        /// <summary>
        /// The image preview for the current template
        /// </summary>
        public Byte[] TemplateImageFile { get; set; }

        /// <summary>
        /// The file name of the image preview for the current template
        /// </summary>
        public String TemplateImageFileName { get; set; }
    }
}
