using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core.Framework.TimerJobs;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.Utilities;

namespace OfficeDevPnP.PartnerPack.WebJobs.ExternalUsersWebJob
{
	class ValidateExternalUsersTimerJob : TimerJob
	{
		private const string PNPCHECKDATEPROPERTYBAGKEY = "_PnP_PartnerPack_ExternalUsersCheckDate";

		private const string CSSSTYLE = @"<style type=""text/css"">
tg {border-collapse: collapse;	border-spacing: 0;}
.tg td {font-family: Arial, sans-serif;font-size: 14px;padding: 10px 5px;border-style: solid;border-width: 1px;overflow: hidden;word-break: normal;}
.tg th {font-family: Arial, sans-serif;font-size: 14px;font-weight: normal;padding: 10px 5px;border-style: solid;border-width: 1px;overflow: hidden;word-break: normal;text-align: left;background-color: #3166ff;color: #ffffff;}
</style>";

		public ValidateExternalUsersTimerJob() : base("Validate External Users Timer Job")
		{
			TimerJobRun += ValidateExternalUsersTimerJob_TimerJobRun;
		}

		private void ValidateExternalUsersTimerJob_TimerJobRun(object sender, TimerJobRunEventArgs e)
		{
			Console.WriteLine("Starting job");
			var web = e.SiteClientContext.Web;
			Tenant tenant = new Tenant(e.TenantClientContext);

			var siteAdmins = e.SiteClientContext.LoadQuery(web.SiteUsers.Include(u => u.Email).Where(u => u.IsSiteAdmin));
			e.SiteClientContext.ExecuteQueryRetry();

			List<string> adminEmails = new List<string>();

			foreach (var siteAdmin in siteAdmins)
			{
				adminEmails.Add(siteAdmin.Email);
			}

			SiteProperties p = tenant.GetSitePropertiesByUrl(e.SiteClientContext.Url, true);
			var sharingCapability = p.EnsureProperty(s => s.SharingCapability);
			if (sharingCapability != Microsoft.Online.SharePoint.TenantManagement.SharingCapabilities.Disabled)
			{
				DateTime checkDate = DateTime.Now;
				var lastCheckDate = e.WebClientContext.Web.GetPropertyBagValueString(PNPCHECKDATEPROPERTYBAGKEY, string.Empty);
				if (lastCheckDate == string.Empty)
				{
					// new site. Temporary set the check date to less than one Month
					checkDate = checkDate.AddMonths(-2);
				}
				else
				{

					if (!DateTime.TryParse(lastCheckDate, out checkDate))
					{
						// Something went wrong with trying to parse the date in the propertybag. Do the check anyway.
						checkDate = checkDate.AddMonths(-2);
					}
				}
				if (checkDate.AddMonths(1) < DateTime.Now)
				{
					e.SiteClientContext.Web.EnsureProperty(w => w.Url);
					e.WebClientContext.Web.EnsureProperty(w => w.Url);
					EmailProperties mailProps = new EmailProperties();
					mailProps.Subject = "Review required: external users with access to your site";
					StringBuilder bodyBuilder = new StringBuilder();

					bodyBuilder.AppendFormat("<html><head>{0}</head><body style=\"font-family:sans-serif\">", CSSSTYLE);
					bodyBuilder.AppendFormat("<p>Your site with address {0} has one or more external users registered. Please review the following list and take appropriate action if such access is not wanted anymore for a user.</p>", e.SiteClientContext.Web.Url);
					bodyBuilder.Append("<table class=\"tg\"><tr><th>Name</th><th>Invited by</th><th>Created</th><th>Invited As</th><th>Accepted As</th></tr>");

					var externalusers = e.TenantClientContext.Web.GetExternalUsersForSiteTenant(new Uri(e.WebClientContext.Web.Url));
					if (externalusers.Any())
					{
						foreach (var externalUser in externalusers)
						{
							bodyBuilder.AppendFormat("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>", externalUser.DisplayName, externalUser.InvitedBy, externalUser.WhenCreated, externalUser.InvitedAs, externalUser.AcceptedAs);
						}
						bodyBuilder.Append("</table></body></html>");
						mailProps.Body = bodyBuilder.ToString();
						mailProps.To = adminEmails.ToArray();

						Utility.SendEmail(e.SiteClientContext, mailProps);
						e.SiteClientContext.ExecuteQueryRetry();
					}
					e.WebClientContext.Web.SetPropertyBagValue(PNPCHECKDATEPROPERTYBAGKEY, DateTime.Now.ToString());
				}

			}

			Console.WriteLine("Ending job");

		}
	}
}
