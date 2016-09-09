using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Defines the common information for a Provisioning Job
    /// </summary>
    public class ProvisioningJobInformation
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
        /// The .NET Type of the Provisioning Job
        /// </summary>
        public String Type { get; set; }

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

        /// <summary>
        /// Provides the Stream of the Provisioning Job file
        /// </summary>
        public Stream JobFile { get; set; }

        /// <summary>
        /// Defines the date and time when the job was scheduled
        /// </summary>
        public DateTime ScheduledOn { get; set; }
    }
}
