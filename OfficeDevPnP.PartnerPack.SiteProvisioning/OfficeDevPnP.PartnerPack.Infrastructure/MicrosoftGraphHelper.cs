using Microsoft.Graph;
using Microsoft.IdentityModel.Claims;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    public static class MicrosoftGraphHelper
    {
        public static GraphServiceClient GetNewGraphClient(string accessToken = null)
        {
            var client = new GraphServiceClient(
              new DelegateAuthenticationProvider(
                  (requestMessage) =>
                  {
                      if (String.IsNullOrEmpty(accessToken))
                      {
                          // Get back the access token.
                          accessToken = MicrosoftGraphHelper.GetAccessTokenForCurrentUser();
                      }

                      if (!String.IsNullOrEmpty(accessToken))
                      {
                          // Configure the HTTP bearer Authorization Header
                          requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                      }
                      else
                      {
                          throw new Exception("Invalid authorization context");
                      }

                      return (Task.FromResult(0));
                  }
                  )
            );

            return client;
        }

        /// <summary>
        /// This helper method returns and OAuth Access Token for the current user
        /// </summary>
        /// <param name="resourceId">The resourceId for which we are requesting the token</param>
        /// <returns>The OAuth Access Token value</returns>
        public static String GetAccessTokenForCurrentUser(String resourceId = null)
        {
            String accessToken = null;
            if (String.IsNullOrEmpty(resourceId))
            {
                resourceId = MicrosoftGraphConstants.MicrosoftGraphResourceId;
            }

            try
            {
                ClientCredential credential = new ClientCredential(
                    PnPPartnerPackSettings.ClientId,
                    PnPPartnerPackSettings.ClientSecret);
                string signedInUserID = System.Security.Claims.ClaimsPrincipal.Current.FindFirst(
                    ClaimTypes.NameIdentifier).Value;
                string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst(
                    "http://schemas.microsoft.com/identity/claims/tenantid").Value;
                AuthenticationContext authContext = new AuthenticationContext(
                    PnPPartnerPackSettings.AADInstance + tenantId,
                    new SessionADALCache(signedInUserID));

                AuthenticationResult result = authContext.AcquireTokenSilent(
                    resourceId,
                    credential,
                    UserIdentifier.AnyUser);

                accessToken = result.AccessToken;
            }
            catch (AdalException ex)
            {
                if (ex.ErrorCode == "failed_to_acquire_token_silently")
                {
                    // Refresh the access token from scratch
                    ForceOAuthChallenge();
                }
                else
                {
                    // Rethrow the exception
                    throw ex;
                }
            }

            return (accessToken);
        }

        private static void ForceOAuthChallenge()
        {
            HttpContext.Current.GetOwinContext().Authentication.Challenge(
                new AuthenticationProperties
                {
                    RedirectUri = HttpContext.Current.Request.Url.ToString(),
                },
                OpenIdConnectAuthenticationDefaults.AuthenticationType);
        }
    }
}
