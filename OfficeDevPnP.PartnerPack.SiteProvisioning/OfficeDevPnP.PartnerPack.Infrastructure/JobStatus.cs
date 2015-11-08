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
    [Flags]
    public enum ProvisioningJobStatus
    {
        /// <summary>
        /// The Provisioning Job is still pending
        /// </summary>
        Pending = 2,
        /// <summary>
        /// The Provisioning Job failed
        /// </summary>
        Failed = 4,
        /// <summary>
        /// The Provisioning Job has been cancelled
        /// </summary>
        Cancelled = 8,
        /// <summary>
        /// The Provisioning Job is running
        /// </summary>
        Running = 16,
        /// <summary>
        /// The Provisioning Job as been completed
        /// </summary>
        Provisioned = 32,
    }
}
