using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioningWeb.Models
{
    public class CreateSiteCollectionViewModel : CreateSiteViewModel
    {
        [Required(ErrorMessage = "Primary Site Collection Administrator is a required field!")]
        [DisplayName("Primary Site Collection Administrator")]
        public PeoplePickerUser[] PrimarySiteCollectionAdmin { get; set; }

        [Required(ErrorMessage = "Secondary Site Collection Administrator is a required field!")]
        [DisplayName("Secondary Site Collection Administrator")]
        public PeoplePickerUser[] SecondarySiteCollectionAdmin { get; set; }

        [DisplayName("Storage Quota")]
        public Int32 StorageQuota { get; set; }

        [DisplayName("ResourceQuota")]
        public Int32 ResourceQuota { get; set; }
    }
}