using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using OfficeDevPnP.PartnerPack.Setup.ViewModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

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
            Boolean? resourceGroupExists = null;

            try
            {
                await HttpHelper.MakeHeadRequestAsync(
                    $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourcegroups/{resourceGroupName}{apiVersion}",
                    accessToken);

                resourceGroupExists = true;
            }
            catch (ApplicationException ex)
            {
                if (ex.InnerException != null)
                {
                    var statusCode = (ex.InnerException as HttpException)?.GetHttpCode();
                    if (statusCode == 404)
                    {
                        resourceGroupExists = false;
                    }
                }
            }

            if (!resourceGroupExists.Value)
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
            else
            {
                return (resourceGroupExists.Value);
            }
        }

        public static async Task RegisterAzureProvider(String accessToken, Guid subscriptionId, String providerNamespace)
        {
            var jsonProviderRegistration = await HttpHelper.MakePostRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/providers/{providerNamespace}/register{apiVersion}",
                accessToken: accessToken);
        }

        public static async Task CreateServicePlan(String accessToken, Guid subscriptionId, String resourceGroupName, String servicePlanName, String location)
        {
            Boolean? servicePlanExists = null;

            try
            {
                await HttpHelper.MakeGetRequestForStringAsync(
                    $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/serverfarms/{servicePlanName}?api-version=2015-08-01",
                    accessToken);

                servicePlanExists = true;
            }
            catch (ApplicationException ex)
            {
                if (ex.InnerException != null)
                {
                    var statusCode = (ex.InnerException as HttpException)?.GetHttpCode();
                    if (statusCode == 404)
                    {
                        servicePlanExists = false;
                    }
                }
            }

            if (!servicePlanExists.Value)
            {
                var jsonServicePlanCreated = await HttpHelper.MakePutRequestForStringAsync(
                    $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/serverfarms/{servicePlanName}?api-version=2015-08-01",
                    new { Kind = "app", Location = location, Name = servicePlanName, Sku = new { Capacity = 1, Family = "B", Name = "B1", Size = "B1", Tier = "Basic" }, Type = "Microsoft.Web/serverfarms" },
                    "application/json",
                    accessToken);
            }
        }

        public static async Task CreateAppServiceWebSite(String accessToken, Guid subscriptionId, String resourceGroupName, String servicePlanName, String appServiceName, String location, AzureAppServiceSetting[] appSettings)
        {
            Boolean? webSiteExists = null;

            try
            {
                await HttpHelper.MakeGetRequestForStringAsync(
                    $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/sites/{appServiceName}?api-version=2016-08-01",
                    accessToken);

                webSiteExists = true;
            }
            catch (ApplicationException ex)
            {
                if (ex.InnerException != null)
                {
                    var statusCode = (ex.InnerException as HttpException)?.GetHttpCode();
                    if (statusCode == 404)
                    {
                        webSiteExists = false;
                    }
                }
            }

            if (!webSiteExists.Value)
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
        }

        public static async Task UploadCertificateToAzureAppService(String accessToken, Guid subscriptionId, String resourceGroupName, String appServiceName, String location, Byte[] pfxBlob, String certificatePassword)
        {
            var jsonAppServiceCreated = await HttpHelper.MakePutRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Web/certificates/{appServiceName}-pfx?api-version=2016-03-01",
                new
                {
                    Location = location,
                    Properties = new
                    {
                        PfxBlob = pfxBlob,
                        Password = certificatePassword
                    }
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
            Boolean? storageAccountExists = null;

            try
            {
                await HttpHelper.MakeGetRequestForStringAsync(
                    $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Storage/storageAccounts/{storageAccountName.ToLower()}?api-version=2016-12-01",
                    accessToken);

                storageAccountExists = true;
            }
            catch (ApplicationException ex)
            {
                if (ex.InnerException != null)
                {
                    var statusCode = (ex.InnerException as HttpException)?.GetHttpCode();
                    if (statusCode == 404)
                    {
                        storageAccountExists = false;
                    }
                }
            }

            if (!storageAccountExists.Value)
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
            }

            var jsonStorageKeys = await HttpHelper.MakePostRequestForStringAsync(
                $"{AzureManagementApiURI}subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.Storage/storageAccounts/{storageAccountName.ToLower()}/listKeys?api-version=2016-12-01",
                accessToken: accessToken);

            var keys = JsonConvert.DeserializeObject<StorageKeys>(jsonStorageKeys);
            return (keys.keys[0].value);
        }

        public static async Task UploadWebJob(String appServiceName, String username, String password, string jobName, string zipPath, JobType jobType)
        {
            // Prepare the BASIC authentication token
            var encoding = System.Text.Encoding.GetEncoding("ISO-8859-1");
            var token = System.Convert.ToBase64String(encoding.GetBytes($"{username}:{password}"));

            var httpClient = new HttpClient();
            httpClient.BaseAddress = new Uri($"https://{appServiceName}.scm.azurewebsites.net/");
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Basic {token}");
            
            using (StreamReader reader = new StreamReader(zipPath))
            {
                StreamContent streamContent = new StreamContent(reader.BaseStream);
                streamContent.Headers.ContentType = new MediaTypeHeaderValue("application/zip");
                streamContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = $"{jobName}.zip"
                };
                var response = await httpClient.PutAsync($"api/zip/site/wwwroot/App_Data/jobs/{jobType.ToString().ToLower()}/{jobName}/", streamContent);
                var result = await response.Content.ReadAsStringAsync();
                if (response.StatusCode != HttpStatusCode.OK)
                {
                    throw new ApplicationException(result);
                }
            }
        }
    }
}
