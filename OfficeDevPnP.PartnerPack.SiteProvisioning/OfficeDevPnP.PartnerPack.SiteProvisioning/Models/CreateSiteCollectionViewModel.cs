using OfficeDevPnP.PartnerPack.Infrastructure;
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
        private PrincipalsViewModel _primarySiteCollectionAdmin;

        [Required(ErrorMessage = "Primary Site Collection Administrator is a required field")]
        [DisplayName("Primary Site Collection Administrator")]
        [UIHint("Principals")]
        public PrincipalsViewModel PrimarySiteCollectionAdmin
        {
            get
            {
                if (_primarySiteCollectionAdmin == null)
                {
                    _primarySiteCollectionAdmin = new PrincipalsViewModel();
                    _primarySiteCollectionAdmin.MaxSelectableProfilesNumber = 1;
                    _primarySiteCollectionAdmin.IncludeGroups = false;
                }

                return _primarySiteCollectionAdmin;
            }
            set
            {
                _primarySiteCollectionAdmin = value;
            }
        }

        private PrincipalsViewModel _secondarySiteCollectionAdmin;

        [Required(ErrorMessage = "Secondary Site Collection Administrator is a required field")]
        [DisplayName("Secondary Site Collection Administrator")]
        [UIHint("Principals")]
        public PrincipalsViewModel SecondarySiteCollectionAdmin
        {
            get
            {
                if (_secondarySiteCollectionAdmin == null)
                {
                    _secondarySiteCollectionAdmin = new PrincipalsViewModel();
                    _secondarySiteCollectionAdmin.MaxSelectableProfilesNumber = 1;
                    _secondarySiteCollectionAdmin.IncludeGroups = false;
                }

                return _secondarySiteCollectionAdmin;
            }
            set
            {
                _secondarySiteCollectionAdmin = value;
            }
        }

        [DisplayName("Storage Maximum Level (MB)")]
        public Int64 StorageMaximumLevel { get; set; } = 1000;

        [DisplayName("Storage Warning Level (MB)")]
        public Int64 StorageWarningLevel { get; set; } = 900;

        [DisplayName("User Code Maximum Level")]
        public Int64 UserCodeMaximumLevel { get; set; } = 0;

        [DisplayName("User Code Warning Level")]
        public Int64 UserCodeWarningLevel { get; set; } = 0;

        [DisplayName("External Sharing")]
        public Boolean ExternalSharingEnabled { get; set; }

        [DisplayName("Enable Partner Pack Extensions")]
        public Boolean PartnerPackExtensionsEnabled { get; set; }

        [DisplayName("Enable Responsive Design Extensions")]
        public Boolean ResponsiveDesignEnabled { get; set; }

        [DisplayName("Managed Path")]
        [UIHint("ManagedPath")]
        public String ManagedPath { get; set; }
    }
}