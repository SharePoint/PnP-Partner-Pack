using Microsoft.Identity.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    /// <summary>
    /// This class handles the real setup process
    /// </summary>
    public static class SetupManager
    {
        /// <summary>
        /// This method handles the real setup process
        /// </summary>
        /// <returns></returns>
        public static async Task SetupPartnerPackAsync(SetupInformation info)
        {
            #region Acquire the Access Token for Azure AD

            //// ***************************************************************************************
            //// First of all get the access token to create the application in Azure AD
            //// ***************************************************************************************

            //// Prepare the MSAL PublicClientApplication to get the access token
            //String MSAL_ClientID = ConfigurationManager.AppSettings["MSAL:ClientId"];
            //PublicClientApplication clientApplication =
            //    new PublicClientApplication(MSAL_ClientID);

            //// Configure the permissions
            //String[] scopes = {
            //            // "Directory.Read.All",
            //            // "Directory.ReadWrite.All",
            //            "Directory.AccessAsUser.All",
            //            };

            //// Acquire an access token for the given scope.
            //var authenticationResult = clientApplication.AcquireTokenAsync(scopes).GetAwaiter().GetResult();

            //// Get back the access token.
            //var accessToken = authenticationResult.Token;

            #endregion

            #region Create the Infrastructural Site Collection

            await UpdateProgress(info, SetupStep.CreateInfrastructuralSiteCollection, "Creating Infrastructural Site Collection");
            CreateInfrastructuralSiteCollection(info);

            #endregion

            #region Create or manage the X.509 Certificate for the App-Only token

            await UpdateProgress(info, SetupStep.ConfigureX509Certificate, "Configuring X.509 Certificate");

            #endregion

            #region Register the Azure AD Application

            await UpdateProgress(info, SetupStep.RegisterAzureADApplication, "Registering Azure AD Application");

            #endregion

            #region Create the Azure Blob Storage

            await UpdateProgress(info, SetupStep.CreateBlobStorageAccount, "Creating Azure Blob Storage Account");

            #endregion

            #region Create the Azure App Service

            await UpdateProgress(info, SetupStep.CreateAzureAppService, "Creating Azure App Service");

            #endregion

            #region Configure the .config files of the App Service, before uploading files

            await UpdateProgress(info, SetupStep.ConfigureSettings, "Configuring Settings");

            #endregion

            #region Provision the WebJobs

            await UpdateProgress(info, SetupStep.ProvisionWebJobs, "Provisioning WebJobs");

            #endregion

            await UpdateProgress(info, SetupStep.Completed, "Setup Completed");
        }

        private static void CreateInfrastructuralSiteCollection(SetupInformation info)
        {
            Uri infrastructureSiteUri = new Uri(info.InfrastructuralSiteUrl);
            Uri tenantAdminUri = new Uri(infrastructureSiteUri.Scheme + "://" +
                infrastructureSiteUri.Host.Replace(".sharepoint.com", "-admin.sharepoint.com/"));
            var siteUrl = info.InfrastructuralSiteUrl.Substring(info.InfrastructuralSiteUrl.IndexOf("sharepoint.com/") + 14);
            var siteCreated = false;

            AuthenticationManager am = new AuthenticationManager();
            using (var adminContext = am.GetAzureADAccessTokenAuthenticatedContext(
                tenantAdminUri.ToString(), info.Office365AccessToken))
            {
                adminContext.RequestTimeout = Timeout.Infinite;

                var tenant = new Tenant(adminContext);

                // Check if the site already exists
                if (true || !tenant.SiteExists(info.InfrastructuralSiteUrl))
                {
                    // Configure the Site Collection properties
                    SiteEntity newSite = new SiteEntity();
                    newSite.Description = "PnP Partner Pack - Infrastructural Site Collection";
                    newSite.Lcid = (uint)info.InfrastructuralSiteLCID;
                    newSite.Title = newSite.Description;
                    newSite.Url = info.InfrastructuralSiteUrl;
                    newSite.SiteOwnerLogin = info.InfrastructuralSitePrimaryAdmin;
                    newSite.StorageMaximumLevel = 1000;
                    newSite.StorageWarningLevel = 900;
                    newSite.Template = "STS#0";
                    newSite.TimeZoneId = info.InfrastructuralSiteTimeZone;
                    newSite.UserCodeMaximumLevel = 0;
                    newSite.UserCodeWarningLevel = 0;

                    // Create the Site Collection and wait for its creation (we're asynchronous)
                    tenant.CreateSiteCollection(newSite, true, true);

                    Site site = tenant.GetSiteByUrl(siteUrl);
                    Web web = site.RootWeb;

                    adminContext.Load(site, s => s.Url);
                    adminContext.Load(web, w => w.Url);
                    adminContext.ExecuteQueryRetry();

                    // Enable Secondary Site Collection Administrator
                    if (!String.IsNullOrEmpty(info.InfrastructuralSiteSecondaryAdmin))
                    {
                        Microsoft.SharePoint.Client.User secondaryOwner = web.EnsureUser(info.InfrastructuralSiteSecondaryAdmin);
                        secondaryOwner.IsSiteAdmin = true;
                        secondaryOwner.Update();

                        web.SiteUsers.AddUser(secondaryOwner);
                        adminContext.ExecuteQueryRetry();
                    }

                    siteCreated = true;
                }
            }

            if (siteCreated)
            {
                using (ClientContext clientContext = am.GetAzureADAccessTokenAuthenticatedContext(
                    info.InfrastructuralSiteUrl, info.Office365AccessToken))
                {
                    clientContext.RequestTimeout = Timeout.Infinite;

                    Site site = clientContext.Site;
                    Web web = site.RootWeb;

                    clientContext.Load(site, s => s.Url);
                    clientContext.Load(web, w => w.Url);
                    clientContext.ExecuteQueryRetry();

                    // Override settings within templates, before uploading them
                    UpdateProvisioningTemplateParameter("Responsive", "SPO-Responsive.xml",
                        "AzureWebSiteUrl", info.AzureWebAppUrl);
                    UpdateProvisioningTemplateParameter("Overrides", "PnP-Partner-Pack-Overrides.xml",
                        "AzureWebSiteUrl", info.AzureWebAppUrl);

                    // Apply the templates to the target site
                    ApplyProvisioningTemplate(web, "Infrastructure", "PnP-Partner-Pack-Infrastructure-Jobs.xml");
                    ApplyProvisioningTemplate(web, "Infrastructure", "PnP-Partner-Pack-Infrastructure-Templates.xml");
                    ApplyProvisioningTemplate(web, "", "PnP-Partner-Pack-Infrastructure-Contents.xml");
                }
            }
            else
            {
                // TODO: Handle some kind of exception ...
            }
        }

        private static void UpdateProvisioningTemplateParameter(string container, string filename, string parameterName, string parameterValue)
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\..\OfficeDevPnP.PartnerPack.SiteProvisioning\Templates",
                    AppDomain.CurrentDomain.BaseDirectory),
                    container);

            ProvisioningTemplate template = provider.GetTemplate(filename);

            if (template.Parameters.ContainsKey(parameterName))
            {
                template.Parameters[parameterName] = parameterValue;
            }

            provider.SaveAs(template, filename);
        }

        private static void ApplyProvisioningTemplate(Web web, string container, string filename)
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\..\OfficeDevPnP.PartnerPack.SiteProvisioning\Templates",
                    AppDomain.CurrentDomain.BaseDirectory),
                    container);

            ProvisioningTemplate template = provider.GetTemplate(filename);
            template.Connector = provider.Connector;

            ProvisioningTemplateApplyingInformation ptai =
                new ProvisioningTemplateApplyingInformation();

            web.ApplyProvisioningTemplate(template, ptai);
        }

        private static async Task UpdateProgress(SetupInformation info, SetupStep currentStep, String stepDescription)
        {
            info.ViewModel.SetupProgress = (100 / 7) * (Int32)currentStep;
            info.ViewModel.SetupProgressDescription = stepDescription;

            await Task.Delay(2000);
        }
    }

    public enum SetupStep
    {
        CreateInfrastructuralSiteCollection,
        ConfigureX509Certificate,
        RegisterAzureADApplication,
        CreateBlobStorageAccount,
        CreateAzureAppService,
        ConfigureSettings,
        ProvisionWebJobs,
        Completed,
    }
}
