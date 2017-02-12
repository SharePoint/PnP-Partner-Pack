using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using OfficeDevPnP.PartnerPack.Setup.ViewModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public static class AzureManagementUtility
    {
        private static String _accessToken;

        public static String AccessToken => _accessToken;

        private static async Task GetAccessTokenAsync()
        {
            // Resource Uri for Azure Management Services
            string resourceUri = "https://management.azure.com/";

            // ClientID of the Azure AD application
            string clientId = ConfigurationManager.AppSettings["ADAL:ClientId"];

            // Redirect URI for native application
            string redirectUri = "urn:ietf:wg:oauth:2.0:oob";

            // Create an instance of AuthenticationContext to acquire an Azure access token  
            // OAuth2 authority Uri  
            string authorityUri = "https://login.microsoftonline.com/common";
            AuthenticationContext authContext = new AuthenticationContext(authorityUri);

            // Call AcquireToken to get an Azure token from Azure Active Directory token issuance endpoint  
            var platformParams = new PlatformParameters(PromptBehavior.RefreshSession);
            var authResult = await authContext.AcquireTokenAsync(resourceUri, clientId, new Uri(redirectUri), platformParams);

            // Return the Access Token
            _accessToken = authResult.AccessToken;
        }

        public static async Task<Dictionary<Guid, String>> ListSubscriptionsAsync()
        {
            //if (String.IsNullOrEmpty(AccessToken))
            //{
            await GetAccessTokenAsync();
            //}

            // Get the list of subscriptions
            var jsonSubscriptions = HttpHelper.MakeGetRequestForString("https://management.azure.com/subscriptions?api-version=2016-09-01",
                AccessToken);

            // Decode JSON list
            var subscriptions = JsonConvert.DeserializeObject<AzureSubscriptions>(jsonSubscriptions);

            // Return a dictionary of subscriptions with ID and DisplayName
            return (subscriptions.Subscriptions.ToDictionary(i => i.SubscriptionId, i => i.DisplayName));
        }
    }
}
