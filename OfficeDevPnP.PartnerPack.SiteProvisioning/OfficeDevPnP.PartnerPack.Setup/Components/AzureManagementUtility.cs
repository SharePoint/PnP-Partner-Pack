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
        private static String apiVersion = "?api-version=2016-09-01";
        public static String AzureManagementApiURI = "https://management.azure.com/";
        public static String MicrosoftGraphResourceId = "https://graph.microsoft.com/";
        public static String MicrosoftGraphV1BaseUri = "https://graph.microsoft.com/v1.0/";
        public static String MicrosoftGraphBetaBaseUri = "https://graph.microsoft.com/beta/";

        public static async Task<String> GetUserUniqueId(String clientId = null)
        {
            // ClientID of the Azure AD application
            if (String.IsNullOrEmpty(clientId))
            {
                clientId = ConfigurationManager.AppSettings["ADAL:ClientId"];
            }

            // Create an instance of AuthenticationContext to acquire an Azure access token  
            // OAuth2 authority Uri  
            string authorityUri = "https://login.microsoftonline.com/common";
            AuthenticationContext authContext = new AuthenticationContext(authorityUri);

            // Call AcquireTokenSilent to get an Azure token from Azure Active Directory token issuance endpoint  
            var authResult = await authContext.AcquireTokenSilentAsync(MicrosoftGraphResourceId, clientId);

            // Return the user's unique ID
            return (authResult.UserInfo.DisplayableId);
        }

        public static async Task<String> GetAccessTokenAsync(String resourceUri = null, String clientId = null)
        {
            // Resource Uri for Azure Management Services
            if (String.IsNullOrEmpty(resourceUri))
            {
                resourceUri = AzureManagementApiURI;
            }

            // ClientID of the Azure AD application
            if (String.IsNullOrEmpty(clientId))
            {
                clientId = ConfigurationManager.AppSettings["ADAL:ClientId"];
            }

            // Redirect URI for native application
            string redirectUri = "urn:ietf:wg:oauth:2.0:oob";

            // Create an instance of AuthenticationContext to acquire an Azure access token  
            // OAuth2 authority Uri  
            string authorityUri = "https://login.microsoftonline.com/common";
            AuthenticationContext authContext = new AuthenticationContext(authorityUri);

            try
            {
                // Call AcquireToken to get an Azure token from Azure Active Directory token issuance endpoint  
                var platformParams = new PlatformParameters(PromptBehavior.RefreshSession);
                var authResult = await authContext.AcquireTokenAsync(resourceUri, clientId, new Uri(redirectUri), platformParams);

                // Return the Access Token
                return (authResult.AccessToken);
            }
            catch (AdalServiceException ex)
            {
                if (ex.ErrorCode.Equals("authentication_canceled"))
                {
                    return (null);
                }
                else
                {
                    throw (ex);
                }
            }
        }

        public static async Task<String> GetAccessTokenSilentAsync(String resourceUri, String clientId = null)
        {
            // Resource Uri for Azure Management Services
            if (String.IsNullOrEmpty(resourceUri))
            {
                resourceUri = AzureManagementApiURI;
            }

            // ClientID of the Azure AD application
            if (String.IsNullOrEmpty(clientId))
            {
                clientId = ConfigurationManager.AppSettings["ADAL:ClientId"];
            }

            // Create an instance of AuthenticationContext to acquire an Azure access token  
            // OAuth2 authority Uri  
            string authorityUri = "https://login.microsoftonline.com/common";
            AuthenticationContext authContext = new AuthenticationContext(authorityUri);

            // Call AcquireTokenSilent to get an Azure token from Azure Active Directory token issuance endpoint  
            var authResult = await authContext.AcquireTokenSilentAsync(resourceUri, clientId);

            // Return the Access Token
            return (authResult.AccessToken);
        }

        public static async Task<Dictionary<Guid, String>> ListSubscriptionsAsync(String accessToken)
        {
            // Get the list of subscriptions
            var jsonSubscriptions = await HttpHelper.MakeGetRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions{apiVersion}",
                accessToken);

            // Decode JSON list
            var subscriptions = JsonConvert.DeserializeObject<AzureSubscriptions>(jsonSubscriptions);

            // Return a dictionary of subscriptions with ID and DisplayName
            return (subscriptions.Subscriptions.ToDictionary(i => i.SubscriptionId, i => i.DisplayName));
        }

        public static async Task<Dictionary<String, String>> ListLocations(String accessToken, Guid subscriptionId)
        {
            // Get the list of Locations
            var jsonLocations = await HttpHelper.MakeGetRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/locations{apiVersion}",
                accessToken);

            // Decode JSON list
            var locations = JsonConvert.DeserializeObject<AzureLocations>(jsonLocations);

            // Return a dictionary of subscriptions with ID and DisplayName
            return (locations.Locations.ToDictionary(i => i.Name, i => i.DisplayName));
        }

        public static async Task<Boolean> CreateResourceGroup(String accessToken, Guid subscriptionId, String resourceGroupName, String location)
        {
            var jsonResourceGroup = await HttpHelper.MakePutRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourcegroups/{resourceGroupName}{apiVersion}",
                new { Name = resourceGroupName, Location = location },
                "application/json",
                accessToken);

            // Decode JSON list
            var resourceGroup = JsonConvert.DeserializeObject<ResourceGroupCreation>(jsonResourceGroup);

            // Return a dictionary of subscriptions with ID and DisplayName
            return (resourceGroup.properties.provisioningState == "Succeded");
        }

        public static async Task RegisterAzureProvider(String accessToken, Guid subscriptionId, String providerNamespace)
        {
            var jsonProviderRegistration = await HttpHelper.MakePostRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/providers/{providerNamespace}/register{apiVersion}",
                accessToken: accessToken);
        }

        public static async Task CreateServicePlan(String accessToken, Guid subscriptionId, String resourceGroupName, String servicePlanName, String location)
        {
            var jsonServicePlanCreated = await HttpHelper.MakePutRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/serverfarms/{servicePlanName}?api-version=2015-08-01",
                new { Kind = "app", Location = location, Name = servicePlanName, Sku = new { Capacity = 1, Family = "B", Name = "B1", Size = "B1", Tier = "Basic" }, Type = "Microsoft.Web/serverfarms" },
                "application/json",
                accessToken);
        }

        public static async Task CreateAppServiceWebSite(String accessToken, Guid subscriptionId, String resourceGroupName, String servicePlanName, String appServiceName, String location, AzureAppServiceSetting[] appSettings)
        {
            var jsonAppServiceCreated = await HttpHelper.MakePutRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/sites/{appServiceName}?api-version=2016-08-01",
                new
                {
                    Kind = "app",
                    Location = location,
                    Name = appServiceName,
                    @Properties = new
                    {
                        ServerFarmId = $"/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/serverfarms/{servicePlanName}",
                        SiteConfig = new
                        {
                            AppSettings = appSettings
                        }
                    },
                    Type = "Microsoft.Web/sites"
                },
                "application/json",
                accessToken);
        }

        public static async Task UploadCertificateToAzureAppService(String accessToken, Guid subscriptionId, String resourceGroupName, String appServiceName, Byte[] pfxBlob, String certificatePassword)
        {
            var jsonAppServiceCreated = await HttpHelper.MakePutRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/certificates/{appServiceName}-pfx?api-version=2016-03-01",
                new
                {
                    PfxBlob = pfxBlob,
                    Password = certificatePassword
                },
                "application/json",
                accessToken);
        }

        public static async Task<String> GetAppServiceWebSitePublishingSettings(String accessToken, Guid subscriptionId, String resourceGroupName, String appServiceName)
        {
            var xmlPublishingProfile = await HttpHelper.MakePostRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/sites/{appServiceName}/publishxml?api-version=2016-08-01",
                new
                {
                    Format = "WebDeploy",
                },
                "application/json",
                accessToken);

            return (xmlPublishingProfile);
        }

        public static async Task<String> CreateStorageAccount(String accessToken, Guid subscriptionId, String resourceGroupName, String servicePlanName, String storageAccountName, String location)
        {
            var jsonStorageAccountCreated = await HttpHelper.MakePutRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Storage/storageAccounts/{storageAccountName.ToLower()}?api-version=2016-12-01",
                new
                {
                    Kind = "Storage",
                    Location = location,
                    Sku = new
                    {
                        Name = "Standard_GRS"
                    }
                },
                "application/json",
                accessToken);

            // Wait for the Storage Account to be ready
            var succeded = false;
            while (!succeded)
            {
                var jsonStorageAccount = await HttpHelper.MakeGetRequestForStringAsync(
                    $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Storage/storageAccounts/{storageAccountName.ToLower()}?api-version=2016-12-01",
                    accessToken: accessToken);

                succeded = jsonStorageAccount.Contains("Succeeded");

                await Task.Delay(2000);
            }

            var jsonStorageKeys = await HttpHelper.MakePostRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Storage/storageAccounts/{storageAccountName.ToLower()}/listKeys?api-version=2016-12-01",
                accessToken: accessToken);

            var keys = JsonConvert.DeserializeObject<StorageKeys>(jsonStorageKeys);
            return (keys.keys[0].value);
        }
    }
}
