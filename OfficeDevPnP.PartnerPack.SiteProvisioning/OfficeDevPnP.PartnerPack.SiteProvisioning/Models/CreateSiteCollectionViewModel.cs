using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class CreateSiteCollectionViewModel : CreateSiteViewModel
    {
        [Required(ErrorMessage = "Primary Site Collection Administrator is a required field!")]
        [DisplayName("Primary Site Collection Administrator")]
        public PeoplePickerUser[] PrimarySiteCollectionAdmin { get; set; }

        [Required(ErrorMessage = "Secondary Site Collection Administrator is a required field!")]
        [DisplayName("Secondary Site Collection Administrator")]
        public PeoplePickerUser[] SecondarySiteCollectionAdmin { get; set; }

        [DisplayName("Storage Maximum Level")]
        public Int64 StorageMaximumLevel { get; set; }

        [DisplayName("Storage Warning Level")]
        public Int64 StorageWarningLevel { get; set; }

        [DisplayName("User Code Maximum Level")]
        public Int64 UserCodeMaximumLevel { get; set; }

        [DisplayName("User Code Warning Level")]
        public Int64 UserCodeWarningLevel { get; set; }

        [DisplayName("External Sharing")]
        public Boolean ExternalSharingEnabled { get; set; }

        [UIHint("ManagedPath")]
        public String ManagedPath { get; set; }
    }
}