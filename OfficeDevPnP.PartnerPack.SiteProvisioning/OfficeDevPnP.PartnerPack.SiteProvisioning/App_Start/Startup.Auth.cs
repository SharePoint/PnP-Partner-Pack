using System;
using System.Collections.Generic;
using System.Configuration;
using System.IdentityModel.Claims;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using Owin;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Models;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Components;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning
{
    public partial class Startup
    {
        public void ConfigureAuth(IAppBuilder app)
        {
            
            app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);

            app.UseCookieAuthentication(new CookieAuthenticationOptions());
            //app.UseCookieAuthentication(new CookieAuthenticationOptions {
            //    CookieManager = new Components.SystemWebCookieManager()
            //});

            app.UseOpenIdConnectAuthentication(
                new OpenIdConnectAuthenticationOptions
                {
                    ClientId = MSGraphAPISettings.ClientId,
                    Authority = MSGraphAPISettings.AADInstance + "common",
                    TokenValidationParameters = new System.IdentityModel.Tokens.TokenValidationParameters
                    {
                        // instead of using the default validation (validating against a single issuer value, as we do in line of business apps), 
                        // we inject our own multitenant validation logic
                        ValidateIssuer = false,
                    },
                    Notifications = new OpenIdConnectAuthenticationNotifications()
                    {
                        SecurityTokenValidated = (context) => 
                        {
                            return Task.FromResult(0);
                        },
                        AuthorizationCodeReceived = (context) =>
                        {
                            var code = context.Code;

                            ClientCredential credential = new ClientCredential(
                                MSGraphAPISettings.ClientId,
                                MSGraphAPISettings.ClientSecret);
                            string signedInUserID = context.AuthenticationTicket.Identity.FindFirst(
                                ClaimTypes.NameIdentifier).Value;
                            string tenantId = context.AuthenticationTicket.Identity.FindFirst(
                                "http://schemas.microsoft.com/identity/claims/tenantid").Value;

                            AuthenticationContext authContext = new AuthenticationContext(
                                MSGraphAPISettings.AADInstance + tenantId,
                                new SessionADALCache(signedInUserID));
                            AuthenticationResult result = authContext.AcquireTokenByAuthorizationCode(
                                code, new Uri(HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Path)),
                                credential, MSGraphAPISettings.MicrosoftGraphResourceId);

                            return Task.FromResult(0);
                        },
                        AuthenticationFailed = (context) =>
                        {
                            context.OwinContext.Response.Redirect("/Home/Error?message=" + context.Exception.Message);
                            context.HandleResponse(); // Suppress the exception
                            return Task.FromResult(0);
                        }
                    }
                });

        }
    }
}
