using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines a Site Collection Provisioning Job
    /// </summary>
    public class SiteCollectionProvisioningJob : SiteProvisioningJob
    {
        /// <summary>
        /// Defines the Primary Site Collection Administrator for the Site Collection to provisiong
        /// </summary>
        public String PrimarySiteCollectionAdmin { get; set; }

        /// <summary>
        /// Defines the Secondary Site Collection Administrator for the Site Collection to provisiong
        /// </summary>
        public String SecondarySiteCollectionAdmin { get; set; }

        /// <summary>
        /// Defines the Storage Maximum Level for the Site Collection to provisiong
        /// </summary>
        public Int64 StorageMaximumLevel { get; set; }

        /// <summary>
        /// Defines the Storage Warning Level for the Site Collection to provisiong
        /// </summary>
        public Int64 StorageWarningLevel { get; set; }

        /// <summary>
        /// Defines the User Code Maximum Level for the Site Collection to provisiong
        /// </summary>
        public Int64 UserCodeMaximumLevel { get; set; }

        /// <summary>
        /// Defines the User Code Warning Level for the Site Collection to provisiong
        /// </summary>
        public Int64 UserCodeWarningLevel { get; set; }

        /// <summary>
        /// Defines whether to enable External Sharing for the Site Collection to provisiong
        /// </summary>
        public Boolean ExternalSharingEnabled { get; set; }

        /// <summary>
        /// Defines whether to enable PnP Partner Pack extensions for the Site Collection to provisiong
        /// </summary>
        public Boolean PartnerPackExtensionsEnabled { get; set; }
    }
}
