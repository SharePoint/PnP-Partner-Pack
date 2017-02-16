using CERTENROLLLib;
using Microsoft.Identity.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    /// <summary>
    /// This class handles the real setup process
    /// </summary>
    public static class SetupManager
    {
        private static XNamespace PnPProvisioningTemplateSchema = "http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema";

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
            // await CreateInfrastructuralSiteCollectionAsync(info);

            #endregion

            #region Create or manage the X.509 Certificate for the App-Only token

            await UpdateProgress(info, SetupStep.ConfigureX509Certificate, "Configuring X.509 Certificate");

            if (info.SslCertificateGenerate)
            {
                CreateX509Certificate(info);
            }
            else
            {
                LoadX509Certificate(info);
            }

            info.AzureAppKeyCredential = GetX509CertificateInformation(info);

            #endregion

            #region Register the Azure AD Application

            await UpdateProgress(info, SetupStep.RegisterAzureADApplication, "Registering Azure AD Application");
            await RegisterAzureADApplication(info);

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

        #region Create Infrastructural Site Collection

        private async static Task CreateInfrastructuralSiteCollectionAsync(SetupInformation info)
        {
            Uri infrastructureSiteUri = new Uri(info.InfrastructuralSiteUrl);
            Uri tenantAdminUri = new Uri(infrastructureSiteUri.Scheme + "://" +
                infrastructureSiteUri.Host.Replace(".sharepoint.com", "-admin.sharepoint.com"));
            Uri sharepointUri = new Uri(infrastructureSiteUri.Scheme + "://" +
                infrastructureSiteUri.Host + "/");
            var siteUrl = info.InfrastructuralSiteUrl.Substring(info.InfrastructuralSiteUrl.IndexOf("sharepoint.com/") + 14);
            var siteCreated = false;

            var accessToken = await AzureManagementUtility.GetAccessTokenSilentAsync(
                tenantAdminUri.ToString(), ConfigurationManager.AppSettings["O365:ClientId"]);

            AuthenticationManager am = new AuthenticationManager();
            using (var adminContext = am.GetAzureADAccessTokenAuthenticatedContext(
                tenantAdminUri.ToString(), accessToken))
            {
                adminContext.RequestTimeout = Timeout.Infinite;

                var tenant = new Tenant(adminContext);

                // Check if the site already exists
                var siteAlreadyExists = tenant.SiteExists(info.InfrastructuralSiteUrl);
                if (!siteAlreadyExists)
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
                    tenant.CreateSiteCollection(newSite, true, true, (top) =>
                    {
                        if (top == TenantOperationMessage.CreatingSiteCollection)
                        {
                            var maxProgress = (100 / (Int32)SetupStep.Completed);
                            info.ViewModel.SetupProgress += 1;
                            if (info.ViewModel.SetupProgress >= maxProgress)
                            {
                                info.ViewModel.SetupProgress = maxProgress;
                            }
                            Task.Delay(100);
                        }
                        return (false);
                    });

                    Site site = tenant.GetSiteByUrl(info.InfrastructuralSiteUrl);
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

                if (siteAlreadyExists || siteCreated)
                {
                    accessToken = await AzureManagementUtility.GetAccessTokenSilentAsync(
                        sharepointUri.ToString(), ConfigurationManager.AppSettings["O365:ClientId"]);

                    using (ClientContext clientContext = am.GetAzureADAccessTokenAuthenticatedContext(
                        info.InfrastructuralSiteUrl, accessToken))
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

                        // We to it twice to force the content types, due to a small bug in the provisioning engine
                        ApplyProvisioningTemplate(web, "", "PnP-Partner-Pack-Infrastructure-Contents.xml");
                    }
                }
                else
                {
                    // TODO: Handle some kind of exception ...
                }
            }
        }

        private static void UpdateProvisioningTemplateParameter(string container, string filename, string parameterName, string parameterValue)
        {
            var filePath = String.Format(@"{0}..\..\..\OfficeDevPnP.PartnerPack.SiteProvisioning\Templates{1}{2}\{3}",
                    AppDomain.CurrentDomain.BaseDirectory,
                    !String.IsNullOrEmpty(container) ? @"\" : String.Empty,
                    container,
                    filename);

            if (System.IO.File.Exists(filePath))
            {
                filePath = new System.IO.FileInfo(filePath).FullName;
                XElement templateXml = XElement.Load(filePath);
                var targetParameter = templateXml
                    .Descendants(PnPProvisioningTemplateSchema + "Parameter")
                    .FirstOrDefault(p => p.Attribute("Key").Value == parameterName);

                if (targetParameter != null)
                {
                    targetParameter.Value = parameterValue;
                }

                templateXml.Save(filePath);
            }
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
            if (currentStep == SetupStep.Completed)
            {
                info.ViewModel.SetupProgress = 100;
            }
            else
            {
                info.ViewModel.SetupProgress = (100 / (Int32)SetupStep.Completed) * (Int32)currentStep;
            }
            info.ViewModel.SetupProgressDescription = stepDescription;

            await Task.Delay(2000);
        }

        #endregion

        #region Manage X.509 Certificate

        private static void CreateX509Certificate(SetupInformation info)
        {
            var certificate = CreateSelfSignedCertificate(info.SslCertificateCommonName.ToLower(),
                info.SslCertificateStartDate, info.SslCertificateEndDate, info.SslCertificatePassword);

            SaveCertificateFiles(info, certificate);
        }

        private static void LoadX509Certificate(SetupInformation info)
        {
            var certificate = new X509Certificate2(info.SslCertificateFile);
            info.SslCertificateCommonName = certificate.SubjectName.Name;

            SaveCertificateFiles(info, certificate);
        }

        private static void SaveCertificateFiles(SetupInformation info, X509Certificate2 certificate)
        {
            var basePath = String.Format(@"{0}..\..\..\..\Scripts\", AppDomain.CurrentDomain.BaseDirectory);

            var pfx = certificate.Export(X509ContentType.Pfx, info.SslCertificatePassword);
            System.IO.File.WriteAllBytes($@"{basePath}{info.SslCertificateCommonName}.pfx", pfx);

            var cer = certificate.Export(X509ContentType.Cert);
            System.IO.File.WriteAllBytes($@"{basePath}{info.SslCertificateCommonName}.cer", cer);
        }

        public static X509Certificate2 CreateSelfSignedCertificate(string subjectName, DateTime startDate, DateTime endDate, String password)
        {
            // Create DistinguishedName for subject and issuer
            var name = new CX500DistinguishedName();
            name.Encode("CN=" + subjectName, X500NameFlags.XCN_CERT_NAME_STR_NONE);

            // Create a new Private Key for the certificate
            CX509PrivateKey privateKey = new CX509PrivateKey();
            privateKey.ProviderName = "Microsoft RSA SChannel Cryptographic Provider";
            privateKey.KeySpec = X509KeySpec.XCN_AT_KEYEXCHANGE;
            privateKey.Length = 2048;
            privateKey.SecurityDescriptor = "D:PAI(A;;0xd01f01ff;;;SY)(A;;0xd01f01ff;;;BA)(A;;0x80120089;;;NS)";
            privateKey.MachineContext = true;
            privateKey.ExportPolicy = X509PrivateKeyExportFlags.XCN_NCRYPT_ALLOW_EXPORT_FLAG;
            privateKey.Create();

            // Define the hashing algorithm
            var serverauthoid = new CObjectId();
            serverauthoid.InitializeFromValue("1.3.6.1.5.5.7.3.1"); // Server Authentication
            var ekuoids = new CObjectIds();
            ekuoids.Add(serverauthoid);
            var ekuext = new CX509ExtensionEnhancedKeyUsage();
            ekuext.InitializeEncode(ekuoids);

            // Create the self signing request
            var cert = new CX509CertificateRequestCertificate();
            cert.InitializeFromPrivateKey(X509CertificateEnrollmentContext.ContextMachine, privateKey, String.Empty);
            cert.Subject = name;
            cert.Issuer = cert.Subject;
            cert.NotBefore = startDate;
            cert.NotAfter = endDate;
            cert.X509Extensions.Add((CX509Extension)ekuext);
            cert.Encode();

            // Enroll the certificate
            var enroll = new CX509Enrollment();
            enroll.InitializeFromRequest(cert);
            string certData = enroll.CreateRequest(EncodingType.XCN_CRYPT_STRING_BASE64HEADER);
            enroll.InstallResponse(InstallResponseRestrictionFlags.AllowUntrustedCertificate,
                certData, EncodingType.XCN_CRYPT_STRING_BASE64HEADER, String.Empty);

            var base64encoded = enroll.CreatePFX(password, PFXExportOptions.PFXExportChainWithRoot);

            // Instantiate the target class with the PKCS#12 data
            return new X509Certificate2(
                System.Convert.FromBase64String(base64encoded), password,
                System.Security.Cryptography.X509Certificates.X509KeyStorageFlags.Exportable);
        }

        private static String GetX509CertificateInformation(SetupInformation info)
        {
            var basePath = String.Format(@"{0}..\..\..\..\Scripts\", AppDomain.CurrentDomain.BaseDirectory);

            var certificate = new X509Certificate2();
            certificate.Import($@"{basePath}{info.SslCertificateCommonName}.cer");

            var rawCert = certificate.GetRawCertData();
            var base64Cert = System.Convert.ToBase64String(rawCert);
            var rawCertHash = certificate.GetCertHash();
            var base64CertHash = System.Convert.ToBase64String(rawCertHash);
            var KeyId = System.Guid.NewGuid().ToString();

            var keyCredential =
                "{" +
                    "\"customKeyIdentifier\": \"" + base64CertHash + "\"," +
                    "\"keyId\": \"" + KeyId + "\"," +
                    "\"type\": \"AsymmetricX509Cert\"," +
                    "\"usage\": \"Verify\"," +
                    "\"value\":  \"" + base64Cert + "\"" +
                "}";

            return (keyCredential);
        }

        #endregion

        #region Register Azure AD Application

        private async static Task RegisterAzureADApplication(SetupInformation info)
        {
            // Fix the App URL
            if (!info.AzureWebAppUrl.EndsWith("/"))
            {
                info.AzureWebAppUrl = info.AzureWebAppUrl + "/";
            }

            // Load the App Manifest template
            Stream stream = typeof(SetupManager)
                .Assembly
                .GetManifestResourceStream("OfficeDevPnP.PartnerPack.Setup.Resources.azure-ad-app-manifest.json.txt");

            using (StreamReader sr = new StreamReader(stream))
            {
                // Get the JSON manifest
                var jsonApplication = sr.ReadToEnd();

                var application = JsonConvert.DeserializeObject<AzureAdApplication>(jsonApplication);
                var keyCredential = JsonConvert.DeserializeObject<KeyCredential>(info.AzureAppKeyCredential);

                application.displayName = info.ApplicationName;
                application.homepage = info.AzureWebAppUrl;
                application.identifierUris = new List<String>();
                application.identifierUris.Add(info.ApplicationUniqueUri);
                application.keyCredentials = new List<KeyCredential>();
                application.keyCredentials.Add(keyCredential);
                application.replyUrls = new List<String>();
                application.replyUrls.Add(info.AzureWebAppUrl);

                // Get an Access Token to create the application via Microsoft Graph
                var office365AzureADAccessToken = await AzureManagementUtility.GetAccessTokenSilentAsync(
                    AzureManagementUtility.MicrosoftGraphResourceId,
                    ConfigurationManager.AppSettings["O365:ClientId"]);

                // Create the Azure AD Application
                String jsonResponse = await HttpHelper.MakePostRequestForStringAsync(
                    String.Format("{0}applications",
                        AzureManagementUtility.MicrosoftGraphBetaBaseUri),
                    application,
                    "application/json", office365AzureADAccessToken);
            }
        }

        #endregion
    }

    public enum SetupStep
    {
        Starting,
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
