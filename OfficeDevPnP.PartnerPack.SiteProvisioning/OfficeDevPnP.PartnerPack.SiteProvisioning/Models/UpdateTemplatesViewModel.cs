using OfficeDevPnP.PartnerPack.Infrastructure;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    public class RefreshSitesViewModel : JobViewModel
    {
        /// <summary>
        /// Allows to determine the status of the Refresh Sites job
        /// </summary>
        public RefreshJobStatus Status { get; set; }

        /// <summary>
        /// The error message of the last job run, in case of any error
        /// </summary>
        public String ErrorMessage { get; set; }
    }

    /// <summary>
    /// Defines the status of the Refresh Sites job
    /// </summary>
    public enum RefreshJobStatus
    {
        /// <summary>
        /// The job is not running
        /// </summary>
        Idle,
        /// <summary>
        /// The job is scheduled
        /// </summary>
        Scheduled,
        /// <summary>
        /// The job is running
        /// </summary>
        Running,
        /// <summary>
        /// The job failed to run
        /// </summary>
        Failed,
    }
}