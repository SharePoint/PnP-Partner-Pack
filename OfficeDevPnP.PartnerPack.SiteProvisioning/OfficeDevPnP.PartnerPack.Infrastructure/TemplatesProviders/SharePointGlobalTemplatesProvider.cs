using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.PartnerPack.Infrastructure.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders
{
    /// <summary>
    /// Implements the Templates Provider that use the local tenant-scoped Site Collection as a repository
    /// </summary>
    public class SharePointGlobalTemplatesProvider : SharePointBaseTemplatesProvider
    {
        public override string DisplayName
        {
            get { return ("Global Tenant"); }
        }

        public SharePointGlobalTemplatesProvider()
            : base(PnPPartnerPackSettings.InfrastructureSiteUrl)
        {

        }

        public override ProvisioningTemplate GetProvisioningTemplate(string templateUri)
        {
            return (base.GetProvisioningTemplate(templateUri));
        }

        public override  ProvisioningTemplateInformation[] SearchProvisioningTemplates(string searchText, TargetPlatform platforms, TargetScope scope)
        {
            return (base.SearchProvisioningTemplates(searchText, platforms, scope));
        }
    }
}
