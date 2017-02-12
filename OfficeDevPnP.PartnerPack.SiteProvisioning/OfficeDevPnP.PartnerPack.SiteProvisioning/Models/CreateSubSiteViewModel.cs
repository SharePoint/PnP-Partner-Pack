using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class CreateSubSiteViewModel : CreateSiteViewModel
    {
        [Display(Name = "SiteInheritsPermissions", ResourceType = typeof(OfficeDevPnP.PartnerPack.Localization.Resource))]
        public bool InheritPermissions { get; set; }
    }
}