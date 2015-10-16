using OfficeDevPnP.Core.Framework.TimerJobs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.ProvisioningJob
{
    public class PnPPartnerPackProvisioningJob : TimerJob
    {
        public PnPPartnerPackProvisioningJob(): base("PnP Partner Pack Provisioning Job")
        {
            TimerJobRun += ExecuteProvisioningJobs;
        }
        private void ExecuteProvisioningJobs(object sender, TimerJobRunEventArgs e)
        {
            Console.WriteLine("Starting job");
            var web = e.SiteClientContext.Web;
            Console.WriteLine("Ending job");
        }
    }
}
