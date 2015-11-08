using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class MyProvisionedSitesViewModel
    {
        public ProvisioningJob[] PersonalJobs { get; set; }
    }
}