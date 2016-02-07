using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Components
{
    public static class MSGraphAPISettings
    {
        public static string ClientId = ConfigurationManager.AppSettings["ida:ClientId"];
        public static string ClientSecret = ConfigurationManager.AppSettings["ida:ClientSecret"];
        public static string AADInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        public static string MicrosoftGraphResourceId = "https://graph.microsoft.com";
    }
}