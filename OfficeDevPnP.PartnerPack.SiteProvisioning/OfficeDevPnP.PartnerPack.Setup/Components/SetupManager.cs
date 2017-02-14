using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
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
