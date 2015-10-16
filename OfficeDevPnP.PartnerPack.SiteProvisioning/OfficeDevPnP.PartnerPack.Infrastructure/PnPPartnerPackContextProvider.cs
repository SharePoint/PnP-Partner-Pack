using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Web;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// This type provides an easy way to create a SharePoint CSOM ClientContext instance
    /// </summary>
    public static class PnPPartnerPackContextProvider
    {
        public static ClientContext GetAppOnlyTenantLevelClientContext()
        {
            Uri infrastructureSiteUri = new Uri(PnPPartnerPackSettings.InfrastructureSiteUrl);
            Uri tenantAdminUri = new Uri(infrastructureSiteUri.Scheme + "://" + 
                infrastructureSiteUri.Host.Replace(".sharepoint.com", "-admin.sharepoint.com"));

            return (PnPPartnerPackContextProvider.GetAppOnlyClientContext(tenantAdminUri.ToString()));
        }

        public static ClientContext GetAppOnlyClientContext(String siteUrl)
        {
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            AuthenticationManager authManager = new AuthenticationManager();
            ClientContext context = authManager.GetAzureADAppOnlyAuthenticatedContext(
                siteUrl,
                PnPPartnerPackSettings.ClientId,
                tenantID,
                PnPPartnerPackSettings.AppOnlyCertificate);

            return (context);
        }

        public static ClientContext GetWebApplicationClientContext(String siteUrl, TokenCache tokenCache = null)
        {
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            AuthenticationManager authManager = new AuthenticationManager();
            ClientContext context = authManager.GetAzureADWebApplicationAuthenticatedContext(
                siteUrl,
                (s) => GetTokenForApplication(s, tokenCache));

            return (context);
        }

        private static String GetTokenForApplication(String serviceUri, TokenCache tokenCache)
        {
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            // Get a token for the Graph without triggering any user interaction (from the cache, via multi-resource refresh token, etc)
            ClientCredential clientcred = new ClientCredential(
                PnPPartnerPackSettings.ClientId, 
                PnPPartnerPackSettings.ClientSecret);

            // Initialize AuthenticationContext with the token cache of the currently signed in user, as kept in the app's database
            AuthenticationContext authenticationContext = new AuthenticationContext(PnPPartnerPackSettings.AADInstance + tenantID, tokenCache);

            // Get the Access Token
            AuthenticationResult authenticationResult = authenticationContext.AcquireTokenSilent(
                serviceUri.ToString(), 
                clientcred,
                new UserIdentifier(userObjectID, UserIdentifierType.UniqueId));

            return authenticationResult.AccessToken;
        }
    }
}