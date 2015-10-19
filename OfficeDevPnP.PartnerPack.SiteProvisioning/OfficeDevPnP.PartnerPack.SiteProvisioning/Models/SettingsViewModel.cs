using OfficeDevPnP.PartnerPack.Infrastructure;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class SettingsViewModel
    {
        /// <summary>
        /// The Site Collections in the current Tenant
        /// </summary>
        [DisplayName("Select a Site Collection")]
        public String SelectedSiteCollection { get; set; }

        public SiteCollectionSettings[] SiteCollections { get; set; }
    }

    //public class SiteCollectionItem
    //{
    //    /// <summary>
    //    /// The Title of a Site Collection
    //    /// </summary>
    //    public String Title { get; set; }

    //    /// <summary>
    //    /// The URL of a Site Collection
    //    /// </summary>
    //    public String Url { get; set; }

    //    /// <summary>
    //    /// Defines whether the PnP Partner Pack is enabled or not on the target Site Collection
    //    /// </summary>
    //    public Boolean PnPPartnerPackEnabled { get; set; }
    //}
}