using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core.Framework.TimerJobs;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.Utilities;

namespace OfficeDevPnP.PartnerPack.WebJobs.CheckAdminsWebJob
{
    class EnforceTwoAdministratorsTimerJob : TimerJob
    {
        public EnforceTwoAdministratorsTimerJob() : base("Enforce Two Administrators Job")
        {
            TimerJobRun += EnforceTwoAdministratorsTimerJob_TimerJobRun;
        }

        private void EnforceTwoAdministratorsTimerJob_TimerJobRun(object sender, TimerJobRunEventArgs e)
        {
            Console.WriteLine("Starting job");
            var web = e.SiteClientContext.Web;

            var siteUsers = e.SiteClientContext.LoadQuery(web.SiteUsers.Include(u => u.Email).Where(u => u.IsSiteAdmin));
            e.SiteClientContext.ExecuteQueryRetry();

            if (siteUsers.Count() < 2)
            {
                Console.WriteLine("Site found");
                if (!web.IsPropertyAvailable("Url"))
                {
                    e.SiteClientContext.Load(web, w => w.Url);
                    e.SiteClientContext.ExecuteQueryRetry();
                }
                var adminUser = siteUsers.FirstOrDefault();

                EmailProperties mailProps = new EmailProperties();
                mailProps.Subject = "Action required: assign an additional site administrator to your site";
                mailProps.Body = string.Format("Your site with address {0} has only one site administrator defined, you. Please assign an additional sign administrator.", e.SiteClientContext.Web.Url);
                mailProps.To = new[] { adminUser.Email };
                Utility.SendEmail(e.SiteClientContext, mailProps);
                e.SiteClientContext.ExecuteQueryRetry();
            }
            Console.WriteLine("Ending job");

        }
    }
}
