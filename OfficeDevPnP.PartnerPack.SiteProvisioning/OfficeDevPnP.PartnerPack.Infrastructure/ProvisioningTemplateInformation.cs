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
        /// Defines the Scope of the Provisioning Template
        /// </summary>
        public TemplateScope Scope { get; set; }

        /// <summary>
        /// Defines the URL of the source Site from which the Provisioning Template has been generated
        /// </summary>
        public String TemplateSourceUrl { get; set; }
    }
}
