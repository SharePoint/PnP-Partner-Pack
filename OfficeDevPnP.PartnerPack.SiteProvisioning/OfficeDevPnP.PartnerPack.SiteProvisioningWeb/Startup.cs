using System;
using System.Threading.Tasks;
using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(OfficeDevPnP.PartnerPack.SiteProvisioningWeb.Startup))]

namespace OfficeDevPnP.PartnerPack.SiteProvisioningWeb
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
