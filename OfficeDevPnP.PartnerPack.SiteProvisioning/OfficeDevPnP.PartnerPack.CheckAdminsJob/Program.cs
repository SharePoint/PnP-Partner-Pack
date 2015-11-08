using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using OfficeDevPnP.Core.Framework.TimerJobs;
using OfficeDevPnP.Core;
using System.Configuration;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using OfficeDevPnP.PartnerPack.Infrastructure;

namespace OfficeDevPnP.PartnerPack.CheckAdminsJob
{
    // To learn more about Microsoft Azure WebJobs SDK, please see http://go.microsoft.com/fwlink/?LinkID=320976
    class Program
    {
        static void Main()
        {
            var job = new EnforceTwoAdministratorsTimerJob();

            var provisioningJobs = ProvisioningRepositoryFactory.Current.GetTypedProvisioningJobs<SiteCollectionProvisioningJob>(ProvisioningJobStatus.Provisioned);

            foreach (SiteCollectionProvisioningJob provisioningJob in provisioningJobs)
            {
                var url = PnPPartnerPackSettings.InfrastructureSiteUrl.Substring(0, PnPPartnerPackSettings.InfrastructureSiteUrl.IndexOf(".com/") + 4) + provisioningJob.RelativeUrl;
                job.AddSite(url);
            }

            job.UseAzureADAppOnlyAuthentication(
                        PnPPartnerPackSettings.ClientId,
                        PnPPartnerPackSettings.Tenant,
                        PnPPartnerPackSettings.AppOnlyCertificate);

            job.Run();

        }
    }
}
