using Newtonsoft.Json;
using OfficeDevPnP.PartnerPack.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Components
{
    public class UserUtility
    {
        public static LightGraphUser GetUser(string path)
        {
            String jsonResponse = HttpHelper.MakeGetRequestForString(
                String.Format("{0}{1}",
                    MicrosoftGraphConstants.MicrosoftGraphV1BaseUri, path), 
                MicrosoftGraphHelper.GetAccessTokenForCurrentUser(MicrosoftGraphConstants.MicrosoftGraphResourceId));

            if (jsonResponse != null)
            {
                var user = JsonConvert.DeserializeObject<LightGraphUser>(jsonResponse);
                return (user);
            }
            else
            {
                return (null);
            }
        }

        public static LightGraphUser GetCurrentUser()
        {
            return GetUser("me");
        }
    }
}