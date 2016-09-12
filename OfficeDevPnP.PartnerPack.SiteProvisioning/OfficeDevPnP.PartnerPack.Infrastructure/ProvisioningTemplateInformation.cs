using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Defines a Provisioning Template stored in a Provisioning Repository
    /// </summary>
    public class ProvisioningTemplateInformation
    {
        /// <summary>
        /// Defines the URI of the template file
        /// </summary>
        public String TemplateFileUri { get; set; }

        /// <summary>
        /// Defines the target Scope of the Provisioning Template
        /// </summary>
        public TargetScope Scope { get; set; }

        /// <summary>
        /// Defines the target Platforms of the Provisioning Template
        /// </summary>
        public TargetPlatform Platforms { get; set; }

        /// <summary>
        /// Defines the URL of the source Site from which the Provisioning Template has been generated, if any
        /// </summary>
        public String TemplateSourceUrl { get; set; }

        /// <summary>
        /// Defines the URL of the source Site from which the Provisioning Template has been generated, if any
        /// </summary>
        public String TemplateImageUrl { get; set; }

        /// <summary>
        /// Defines the Display Name of the Provisioning Template
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// Defines the Description of the Provisioning Template
        /// </summary>
        public String Description { get; set; }
    }
}
