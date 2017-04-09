using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    public class RefreshSingleSiteJobHandler : ProvisioningJobHandler
    {
        protected override void RunJobInternal(ProvisioningJob job)
        {
            RefreshSingleSiteJob updateJob = job as RefreshSingleSiteJob;
            if (updateJob == null)
            {
                throw new ArgumentException("Invalid job type for RefreshSingleSiteJobHandler.");
            }

            UpdateTemplates(updateJob);
        }

        private void UpdateTemplates(RefreshSingleSiteJob job)
        {
            using (var siteContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(job.TargetSiteUrl))
            {
                // Get a reference to the target web
                var targetWeb = siteContext.Site.RootWeb;

                // Update the root web of the site collection
                RefreshSitesJobHandler.UpdateTemplateOnWeb(targetWeb);
            }
        }
    }
}
