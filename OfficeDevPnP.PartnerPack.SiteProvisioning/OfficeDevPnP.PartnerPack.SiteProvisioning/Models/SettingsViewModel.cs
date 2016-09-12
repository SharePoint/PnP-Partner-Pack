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
}