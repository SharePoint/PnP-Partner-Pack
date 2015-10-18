using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Defines the general settings of a Site Collection from a PnP Partner Pack perspective
    /// </summary>
    public class SiteCollectionSettings
    {
        /// <summary>
        /// Defines the URL of the Site Collection
        /// </summary>
        public String Url { get; set; }

        /// <summary>
        /// Defines the Title of the Site Collection
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Declares whether the Site Collection has the PnP Partner Pack extensions provisioned
        /// </summary>
        public bool PnPPartnerPackEnabled { get; set; }
    }
}
