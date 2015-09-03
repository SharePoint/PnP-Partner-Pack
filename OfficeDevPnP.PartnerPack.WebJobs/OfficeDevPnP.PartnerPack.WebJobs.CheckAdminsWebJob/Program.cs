using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using OfficeDevPnP.Core.Framework.TimerJobs;
using OfficeDevPnP.Core;
using System.Configuration;

namespace OfficeDevPnP.PartnerPack.WebJobs.CheckAdminsWebJob
{
    // To learn more about Microsoft Azure WebJobs SDK, please see http://go.microsoft.com/fwlink/?LinkID=320976
    class Program
    {
        // Please set the following connection strings in app.config for this WebJob to run:
        // AzureWebJobsDashboard and AzureWebJobsStorage
        static void Main()
        {
            var job = new EnforceTwoAdministratorsTimerJob();

            job.AddSite("https://erwinmcm.sharepoint.com/sites/test");

            //   job.SetEnumerationCredentials("erwinmcm");
            job.UseAzureADAppOnlyAuthentication(
                ConfigurationManager.AppSettings["ClientId"],
                ConfigurationManager.AppSettings["AzureTenant"],
                ConfigurationManager.AppSettings["CertificatePath"],
                ConfigurationManager.ConnectionStrings["CertificatePassword"].ConnectionString);

            job.Run();

        }
    }
}
