using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioningWeb.Models
{
    public class CreateSubSiteViewModel : CreateSiteViewModel
    {
        [DisplayName("Inherit Permissions from Site Collection")]
        public bool InheritPermissions { get; set; }
    }
}