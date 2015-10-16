using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines the base abstract class for any Site Provisioning Job
    /// </summary>
    public abstract class ProvisioningJob
    {
        /// <summary>
        /// The ID of the Provisioning Job
        /// </summary>
        public Guid JobId { get; set; }

        /// <summary>
        /// The descriptive Title of the Provisioning Job
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// The Owner (creator) of the Provisioning Job
        /// </summary>
        public String Owner { get; set; }

        /// <summary>
        /// Defines the Status of the Provisioning Job
        /// </summary>
        public ProvisioningJobStatus Status { get; set; }

        /// <summary>
        /// Defines the Error Message of the Provisioning Job, if any
        /// </summary>
        public String ErrorMessage { get; set; }
    }
}
