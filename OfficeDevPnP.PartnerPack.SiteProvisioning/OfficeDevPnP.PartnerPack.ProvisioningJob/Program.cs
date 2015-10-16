using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using OfficeDevPnP.PartnerPack.Infrastructure;

namespace OfficeDevPnP.PartnerPack.ProvisioningJob
{
    class Program
    {
        static void Main()
        {
            var job = new PnPPartnerPackProvisioningJob();

            job.AddSite(PnPPartnerPackSettings.InfrastructureSiteUrl);

            //job.
            //job.UseAzureADAppOnlyAuthentication(
            //    PnPPartnerPackSettings.ClientId,
            //    PnPPartnerPackSettings.Tenant,
            //    ConfigurationManager.AppSettings["CertificatePath"],
            //    ConfigurationManager.ConnectionStrings["CertificatePassword"].ConnectionString);

            job.Run();
        }
    }
}
