using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders
{
    internal class SharePointSearchCacheKey
    {
        public String TemplatesProviderTypeName { get; set; }

        public String SearchText { get; set; }

        public TargetPlatform Platforms { get; set; }

        public TargetScope Scope { get; set; }
    }
}
