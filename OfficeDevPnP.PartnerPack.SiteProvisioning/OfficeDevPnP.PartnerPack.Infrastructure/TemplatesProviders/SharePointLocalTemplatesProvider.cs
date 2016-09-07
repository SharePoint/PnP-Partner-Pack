using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders
{
    /// <summary>
    /// Implements the Templates Provider that use the local Site Collection as a repository
    /// </summary>
    public class SharePointLocalTemplatesProvider : SharePointBaseTemplatesProvider
    {
        public override string DisplayName
        {
            get { return ("Current Site Collection"); }
        }

        public SharePointLocalTemplatesProvider()
            : base(PnPPartnerPackSettings.ParentSiteUrl)
        {

        }

        // NOTE: Use the current context to determine the URL of the current Site Collection

        public override ProvisioningTemplate GetProvisioningTemplate(string templateUri)
        {
            return (base.GetProvisioningTemplate(templateUri));
        }

        public override ProvisioningTemplateInformation[] SearchProvisioningTemplates(string searchText, TargetPlatform platforms, TargetScope scope)
        {
            return (base.SearchProvisioningTemplates(searchText, platforms, scope));
        }

    }
}
