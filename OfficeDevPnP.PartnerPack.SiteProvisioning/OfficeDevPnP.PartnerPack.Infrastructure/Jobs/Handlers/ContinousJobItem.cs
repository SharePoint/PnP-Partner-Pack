using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    /// <summary>
    /// Defines an item to enqueue on the Azure Storage Queue targeting a job to execute in Continous mode
    /// </summary>
    public class ContinousJobItem
    {
        /// <summary>
        /// The GUID of the Job to execute
        /// </summary>
        public Guid JobId { get; set; }
    }
}
