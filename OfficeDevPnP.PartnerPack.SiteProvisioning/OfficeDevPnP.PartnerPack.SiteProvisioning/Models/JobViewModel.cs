using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    /// <summary>
    /// Basic abstract class that defines a ViewModel for a generic job
    /// </summary>
    public abstract class JobViewModel
    {
        public Guid? JobId { get; set; }
    }
}