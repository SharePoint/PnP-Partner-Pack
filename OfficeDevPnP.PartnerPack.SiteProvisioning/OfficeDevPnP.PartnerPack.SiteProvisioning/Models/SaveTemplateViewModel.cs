using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class SaveTemplateViewModel : JobViewModel
    {
        [Required(ErrorMessage = "Title is a required field")]
        [DisplayName("Title")]
        public String Title { get; set; }

        [DisplayName("Description")]
        [DataType(DataType.MultilineText)]
        [UIHint("Multilines")]
        public String Description { get; set; }

        [Required(ErrorMessage = "Template File Name is a required field")]
        [DisplayName("Template File Name (without extension)")]
        public String FileName { get; set; }

        [DisplayName("Include All Term Groups")]
        public Boolean IncludeAllTermGroups { get; set; }

        [DisplayName("Include Site Collection Term Group")]
        public Boolean IncludeSiteCollectionTermGroup { get; set; }

        [DisplayName("Include Search Configuration")]
        public Boolean IncludeSearchConfiguration { get; set; } = false;

        [DisplayName("Include Site Groups")]
        public Boolean IncludeSiteGroups { get; set; }

        [DisplayName("Persist Composed Look Files")]
        public Boolean PersistComposedLookFiles { get; set; }

        [DisplayName("Target Location for Template")]
        [UIHint("TemplateLocation")]
        public String Location { get; set; }

        public String SourceSiteUrl { get; set; }
        
        public Guid ProvisioningTemplateId { get; set; }
    }
}