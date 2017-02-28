using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Resources;
using System.Text;
using System.Web.Http;
using System.Web.Mvc;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Controllers
{
    public class ResourcesController : Controller
    {
        [OutputCache(Duration = 36000, VaryByParam = "lang")]
        // Duration can be many hours as embedded resources cannot change without recompiling.
        // For the clients, I think 10 hours is good.
        public JavaScriptResult Index(string lang)
        {
            var culture = new CultureInfo(lang);
            return BuildJavaScriptResources(culture);
        }

        public static JavaScriptResult BuildJavaScriptResources(CultureInfo culture)
        {
            var resourceObjectName = "__pnpPartnerPackResources";

            var resourceSet = Localization.Resource.ResourceManager.GetResourceSet(culture, false, true);

            var sb = new StringBuilder();

            sb.AppendFormat("var {0}={};", resourceObjectName);

            var enumerator = resourceSet.GetEnumerator();
            while (enumerator.MoveNext())
            {
                sb.AppendFormat("{0}.{1}='{2}';", resourceObjectName, enumerator.Key,
                    System.Web.HttpUtility.JavaScriptStringEncode(enumerator.Value.ToString()));
            }

            return new JavaScriptResult()
            {
                Script = sb.ToString()
            };
        }
    }
}