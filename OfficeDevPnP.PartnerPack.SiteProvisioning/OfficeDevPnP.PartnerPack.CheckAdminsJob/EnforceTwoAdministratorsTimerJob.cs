using OfficeDevPnP.Core.Framework.TimerJobs;
using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using Microsoft.SharePoint.Client.Utilities;
using System.Text;
using System.Collections.Generic;

namespace OfficeDevPnP.PartnerPack.CheckAdminsJob
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
                StringBuilder bodyBuilder = new StringBuilder();
                bodyBuilder.Append("<html><body style=\"font-family:sans-serif\">");
                bodyBuilder.AppendFormat("<p>Your site with address <a href=\"{0}\">{0}</a> has only one site administrator defined: you. Please assign an additional site administrator.</p>", e.SiteClientContext.Web.Url);
                bodyBuilder.AppendFormat("<p>Click here to <a href=\"{0}/_layouts/mngsiteadmin.aspx\">assign an additional site collection adminstrator.</a></p>", e.SiteClientContext.Web.Url);
                bodyBuilder.Append("</body></html>");
                mailProps.Body = bodyBuilder.ToString();
                mailProps.To = new[] { adminUser.Email };
                Utility.SendEmail(e.SiteClientContext, mailProps);
                e.SiteClientContext.ExecuteQueryRetry();
            }
            Console.WriteLine("Ending job");

        }
    }
}
