using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Contains a bunch of constanta strings and values for the Microsoft Graph
    /// </summary>
    public static class MicrosoftGraphConstants
    {
        public static String MicrosoftGraphV1BaseUri = "https://graph.microsoft.com/v1.0/";
        public static String MicrosoftGraphBetaBaseUri = "https://graph.microsoft.com/beta/";
        public static String MicrosoftGraphResourceId = "https://graph.microsoft.com";

        public static String GlobalTenantAdminRole = "Company Administrator";
        public static String GlobalSPOAdminRole = "SharePoint Service Administrator";
    }
}
