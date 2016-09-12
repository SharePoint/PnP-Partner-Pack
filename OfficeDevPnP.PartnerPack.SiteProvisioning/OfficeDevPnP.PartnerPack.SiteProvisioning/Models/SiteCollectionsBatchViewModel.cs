using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class SiteCollectionsBatchViewModel : JobViewModel
    {
        public BatchStep Step { get; set; }

        public batches Sites { get; set; }

        public String SitesJson { get; set; }
    }

    public enum BatchStep
    {
        BatchStartup,
        BatchFileUploaded,
        BatchScheduled,
    }
}