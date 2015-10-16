using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// The status of a Provisioning Job
    /// </summary>
    public enum JobStatus
    {
        /// <summary>
        /// The Provisioning Job is still pending
        /// </summary>
        Pending,
        /// <summary>
        /// The Provisioning Job failed
        /// </summary>
        Failed,
        /// <summary>
        /// The Provisioning Job has been cancelled
        /// </summary>
        Cancelled,
        /// <summary>
        /// The Provisioning Job as been completed
        /// </summary>
        Provisioned,
    }
}
