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
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Resources;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
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
            #region Create the Infrastructural Site Collection

            await UpdateProgress(info, SetupStep.CreateInfrastructuralSiteCollection, "Creating Infrastructural Site Collection");
            await CreateInfrastructuralSiteCollectionAsync(info);

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

            info.SslCertificateThumbprint = GetX509CertificateThumbprint(info);
            info.AzureAppKeyCredential = GetX509CertificateInformation(info);

            #endregion

            #region Register the Azure AD Application

            await UpdateProgress(info, SetupStep.RegisterAzureADApplication, "Registering Azure AD Application");
            await RegisterAzureADApplication(info);

            #endregion

            #region Create the Azure Resource Group

            await UpdateProgress(info, SetupStep.CreateResourceGroup, "Creating Azure Resource Group");
            await CreateAzureResourceGroup(info);

            #endregion

            #region Create the Azure Blob Storage

            await UpdateProgress(info, SetupStep.CreateBlobStorageAccount, "Creating Azure Blob Storage Account");
            await CreateAzureStorageAccount(info);

            #endregion

            #region Create the Azure App Service

            await UpdateProgress(info, SetupStep.CreateAzureAppService, "Creating Azure App Service");
            await CreateAzureAppService(info);

            #endregion

            #region Configure the .config files of the App Service, before uploading files

            await UpdateProgress(info, SetupStep.ConfigureSettings, "Configuring Settings");
            await ConfigureSettings(info);

            #endregion

            #region Provision the Web Site

            await UpdateProgress(info, SetupStep.ProvisionWebSite, "Provisioning Azure Web Site");
            await BuildAndDeployWebSite(info);

            #endregion

            #region Provision the WebJobs

            await UpdateProgress(info, SetupStep.ProvisionWebJobs, "Provisioning Azure WebJobs");
            await BuildAndDeployJobs(info);

            #endregion

            await UpdateProgress(info, SetupStep.Completed, "Setup Completed");
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
            var siteAlreadyExists = false;

            var accessToken = await AzureManagementUtility.GetAccessTokenSilentAsync(
                tenantAdminUri.ToString(), ConfigurationManager.AppSettings["O365:ClientId"]);

            AuthenticationManager am = new AuthenticationManager();
            using (var adminContext = am.GetAzureADAccessTokenAuthenticatedContext(
                tenantAdminUri.ToString(), accessToken))
            {
                adminContext.RequestTimeout = Timeout.Infinite;

                var tenant = new Tenant(adminContext);

                // Check if the site already exists, and eventually removes it from the Recycle Bin
                if (tenant.CheckIfSiteExists(info.InfrastructuralSiteUrl, "Recycled"))
                {
                    tenant.DeleteSiteCollectionFromRecycleBin(info.InfrastructuralSiteUrl);
                }

                siteAlreadyExists = tenant.SiteExists(info.InfrastructuralSiteUrl);
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
                        }
                        return (false);
                    });
                }
            }

            await Task.Delay(5000);

            using (var adminContext = am.GetAzureADAccessTokenAuthenticatedContext(
                tenantAdminUri.ToString(), accessToken))
            {
                adminContext.RequestTimeout = Timeout.Infinite;

                var tenant = new Tenant(adminContext);
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
            var certificate = new X509Certificate2(info.SslCertificateFile, info.SslCertificatePassword);
            info.AuthenticationCertificate = certificate;
            info.SslCertificateCommonName = certificate.SubjectName.Name;
        }

        private static void SaveCertificateFiles(SetupInformation info, X509Certificate2 certificate)
        {
            info.AuthenticationCertificate = certificate;
            //var basePath = String.Format(@"{0}..\..\..\..\Scripts\", AppDomain.CurrentDomain.BaseDirectory);

            //info.SslCertificateFile = $@"{basePath}{info.SslCertificateCommonName}.pfx";
            //var pfx = certificate.Export(X509ContentType.Pfx, info.SslCertificatePassword);
            //System.IO.File.WriteAllBytes(info.SslCertificateFile, pfx);

            //var cer = certificate.Export(X509ContentType.Cert);
            //System.IO.File.WriteAllBytes($@"{basePath}{info.SslCertificateCommonName}.cer", cer);
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

        private static String GetX509CertificateThumbprint(SetupInformation info)
        {
            var certificate = info.AuthenticationCertificate;
            return (certificate.Thumbprint.ToUpper());
        }

        private static String GetX509CertificateInformation(SetupInformation info)
        {
            // var basePath = String.Format(@"{0}..\..\..\..\Scripts\", AppDomain.CurrentDomain.BaseDirectory);

            var certificate = info.AuthenticationCertificate;
            //var certificate = new X509Certificate2();
            //if (info.SslCertificateGenerate)
            //{
            //    certificate.Import($@"{basePath}{info.SslCertificateCommonName}.cer");
            //}
            //else
            //{
            //    certificate = new X509Certificate2(info.SslCertificateFile, info.SslCertificatePassword);
            //}

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
                .GetManifestResourceStream("OfficeDevPnP.PartnerPack.Setup.Resources.azure-ad-app-manifest.json");

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

                // Generate the Application Shared Secret
                var startDate = DateTime.Now;
                Byte[] bytes = new Byte[32];
                using (var rand = System.Security.Cryptography.RandomNumberGenerator.Create())
                {
                    rand.GetBytes(bytes);
                }
                info.AzureAppSharedSecret = System.Convert.ToBase64String(bytes);
                application.passwordCredentials = new List<object>();
                application.passwordCredentials.Add(new AzureAdApplicationPasswordCredential
                {
                    CustomKeyIdentifier = null,
                    StartDate = startDate.ToString("o"),
                    EndDate = startDate.AddYears(2).ToString("o"),
                    KeyId = Guid.NewGuid().ToString(),
                    Value = info.AzureAppSharedSecret,
                });

                // Get an Access Token to create the application via Microsoft Graph
                var office365AzureADAccessToken = await AzureManagementUtility.GetAccessTokenSilentAsync(
                    AzureManagementUtility.MicrosoftGraphResourceId,
                    ConfigurationManager.AppSettings["O365:ClientId"]);

                var azureAdApplicationCreated = false;

                // Create the Azure AD Application
                try
                {
                    await CreateAzureADApplication(info, application, office365AzureADAccessToken);
                    azureAdApplicationCreated = true;
                }
                catch (ApplicationException ex)
                {
                    var graphError = JsonConvert.DeserializeObject<GraphError>(((HttpException)ex.InnerException).Message);
                    if (graphError != null && graphError.error.code == "Request_BadRequest" &&
                        graphError.error.message.Contains("identifierUris already exists"))
                    {
                        // We need to remove the existing application

                        // Thus, retrieve it
                        String jsonApplications = await HttpHelper.MakeGetRequestForStringAsync(
                            String.Format("{0}applications?$filter=identifierUris/any(c:c+eq+'{1}')",
                                AzureManagementUtility.MicrosoftGraphBetaBaseUri,
                                HttpUtility.UrlEncode(info.ApplicationUniqueUri)),
                            office365AzureADAccessToken);

                        var applications = JsonConvert.DeserializeObject<AzureAdApplications>(jsonApplications);
                        var applicationToUpdate = applications.Applications.FirstOrDefault();
                        if (applicationToUpdate != null)
                        {
                            // Remove it
                            await HttpHelper.MakeDeleteRequestAsync(
                                String.Format("{0}applications/{1}",
                                    AzureManagementUtility.MicrosoftGraphBetaBaseUri,
                                    applicationToUpdate.Id),
                                office365AzureADAccessToken);

                            // And add it again
                            await CreateAzureADApplication(info, application, office365AzureADAccessToken);

                            azureAdApplicationCreated = true;
                        }
                    }
                }

                if (azureAdApplicationCreated)
                {
                    // TODO: We should upload the logo
                    // property mainLogo: stream of the application via PATCH
                }
            }
        }

        private static async Task CreateAzureADApplication(SetupInformation info, AzureAdApplication application, string office365AzureADAccessToken)
        {
            String jsonResponse = await HttpHelper.MakePostRequestForStringAsync(
                String.Format("{0}applications",
                    AzureManagementUtility.MicrosoftGraphBetaBaseUri),
                application,
                "application/json", office365AzureADAccessToken);

            var azureAdApplication = JsonConvert.DeserializeObject<AzureAdApplication>(jsonResponse);
            info.AzureAppClientId = azureAdApplication.AppId.HasValue ? azureAdApplication.AppId.Value : Guid.Empty;
        }

        #endregion

        #region Create the Azure Resources

        private async static Task CreateAzureResourceGroup(SetupInformation info)
        {
            await AzureManagementUtility.RegisterAzureProvider(info.AzureAccessToken,
                info.AzureTargetSubscriptionId,
                "Microsoft.Storage");

            await AzureManagementUtility.RegisterAzureProvider(info.AzureAccessToken,
                info.AzureTargetSubscriptionId,
                "Microsoft.Web");

            info.AzureResourceGroupName = $"{info.AzureAppServiceName}-resource-group";
            info.AzureServicePlanName = $"{info.AzureAppServiceName}-plan";

            await AzureManagementUtility.CreateResourceGroup(info.AzureAccessToken,
                info.AzureTargetSubscriptionId,
                info.AzureResourceGroupName,
                info.AzureLocationId);

            await AzureManagementUtility.CreateServicePlan(info.AzureAccessToken,
                info.AzureTargetSubscriptionId,
                info.AzureResourceGroupName,
                info.AzureServicePlanName,
                info.AzureLocationDisplayName);
        }

        private async static Task CreateAzureStorageAccount(SetupInformation info)
        {
            var key = await AzureManagementUtility.CreateStorageAccount(info.AzureAccessToken,
                info.AzureTargetSubscriptionId,
                info.AzureResourceGroupName,
                info.AzureServicePlanName,
                info.AzureBlobStorageName,
                info.AzureLocationDisplayName);

            info.AzureStorageKey = key;
        }

        private async static Task CreateAzureAppService(SetupInformation info)
        {
            var appSettings = new AzureAppServiceSetting[3];
            appSettings[0] = new AzureAppServiceSetting { Name = "WEBSITE_LOAD_CERTIFICATES", Value = "*" };
            appSettings[1] = new AzureAppServiceSetting { Name = "WEBJOBS_IDLE_TIMEOUT", Value = "10000" };
            appSettings[2] = new AzureAppServiceSetting { Name = "SCM_COMMAND_IDLE_TIMEOUT", Value = "10000" };

            await AzureManagementUtility.CreateAppServiceWebSite(info.AzureAccessToken,
                info.AzureTargetSubscriptionId,
                info.AzureResourceGroupName,
                info.AzureServicePlanName,
                info.AzureAppServiceName,
                info.AzureLocationDisplayName,
                appSettings);

            var certificate = info.AuthenticationCertificate;
            var pfxBlob = certificate.Export(X509ContentType.Pfx, info.SslCertificatePassword);

            await AzureManagementUtility.UploadCertificateToAzureAppService(info.AzureAccessToken,
                info.AzureTargetSubscriptionId,
                info.AzureResourceGroupName,
                info.AzureAppServiceName,
                info.AzureLocationDisplayName,
                pfxBlob,
                info.SslCertificatePassword);

            info.AzureAppPublishingSettings = await AzureManagementUtility.GetAppServiceWebSitePublishingSettings(
                info.AzureAccessToken,
                info.AzureTargetSubscriptionId,
                info.AzureResourceGroupName,
                info.AzureAppServiceName);
        }

        #endregion

        #region Configure Settings

        private async static Task ConfigureSettings(SetupInformation info)
        {
            var basePath = String.Format(@"{0}..\..\..\", AppDomain.CurrentDomain.BaseDirectory);

            var configFiles = new String[5];
            configFiles[0] = (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.CheckAdminsJob\App.config")).FullName;
            configFiles[1] = (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.ExternalUsersJob\App.config")).FullName;
            configFiles[2] = (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.ContinousJob\App.config")).FullName;
            configFiles[3] = (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.ScheduledJob\App.config")).FullName;
            configFiles[4] = (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.SiteProvisioning\Web.config")).FullName;

            var azureStorageConnection = $"DefaultEndpointsProtocol=https;AccountName={info.AzureBlobStorageName};AccountKey={info.AzureStorageKey}";

            foreach (var config in configFiles)
            {
                XElement xmlConfig = XElement.Load(config);

                // Configure Connection Strings
                var connectionStrings = xmlConfig.Element("connectionStrings");
                foreach (var cn in connectionStrings.Elements("add"))
                {
                    if (cn.Attribute("name").Value == "AzureWebJobsDashboard" ||
                        cn.Attribute("name").Value == "AzureWebJobsStorage")
                    {
                        cn.Attribute("connectionString").Value = azureStorageConnection;
                    }
                }

                // Configure AppSettings
                var appSettings = xmlConfig.Element("appSettings");
                foreach (var s in appSettings.Elements("add"))
                {
                    switch (s.Attribute("key").Value)
                    {
                        case "ida:ClientId":
                            s.Attribute("value").Value = info.AzureAppClientId.ToString();
                            break;
                        case "ida:ClientSecret":
                            s.Attribute("value").Value = info.AzureAppSharedSecret;
                            break;
                    }
                }

                // PnP Partner Pack custom settings
                var pnpNamespace = XNamespace.Get("http://schemas.dev.office.com/PnP/2016/08/PnPPartnerPackConfiguration");

                var pnpPartnerPackSettings = xmlConfig.Element(pnpNamespace + "PnPPartnerPackConfiguration");
                if (pnpPartnerPackSettings != null)
                {
                    var tenantSettings = pnpPartnerPackSettings.Element(pnpNamespace + "TenantSettings");

                    if (tenantSettings != null)
                    {
                        if (tenantSettings.Attribute("tenant") != null)
                        {
                            tenantSettings.Attribute("tenant").Value = $"{info.AzureADTenant}.onmicrosoft.com";
                        }
                        if (tenantSettings.Attribute("appOnlyCertificateThumbprint") != null)
                        {
                            tenantSettings.Attribute("appOnlyCertificateThumbprint").Value = info.SslCertificateThumbprint;
                        }
                        if (tenantSettings.Attribute("infrastructureSiteUrl") != null)
                        {
                            tenantSettings.Attribute("infrastructureSiteUrl").Value = info.InfrastructuralSiteUrl;
                        }
                    }
                }
                xmlConfig.Save(config);
            }
        }

        #endregion

        #region Build and Publish Azure Web Site

        private static async Task BuildAndDeployWebSite(SetupInformation info)
        {
            // Get the Project Path
            var basePath = String.Format(@"{0}..\..\..\", AppDomain.CurrentDomain.BaseDirectory);
            var projectPath = (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.SiteProvisioning")).FullName;

            // Save the PublishingSettings file
            var xmlPublishingSettings = XElement.Parse(info.AzureAppPublishingSettings);
            var publishingSettingsPath = projectPath + @"\pnp-partner-pack.publishingSettings";
            xmlPublishingSettings.Save(publishingSettingsPath);

            // Run PowerShell script to build and deploy
            Hashtable packageBuildParameters = new Hashtable();
            packageBuildParameters.Add("ProjectPath", projectPath);
            packageBuildParameters.Add("PublishingSettingsPath", publishingSettingsPath);

            var powerShellScriptPath = (new System.IO.FileInfo($@"{basePath}\OfficeDevPnP.PartnerPack.Setup\Scripts\MsBuildWebSite.ps1")).FullName;
            string packageBuildResult = Run.RunScript(powerShellScriptPath, packageBuildParameters);

            if (packageBuildResult.Contains("Build FAILED"))
            {
                throw new ApplicationException("Failed to build the web project with MSBuild!");
            }
            else if (packageBuildResult.Contains("Missing MSBuild"))
            {
                throw new ApplicationException("Missing MSBuild v. 14.0.25420.1 or higher!");
            }
        }

        #endregion

        #region Build and Publish Azure Web Jobs

        private async static Task BuildAndDeployJobs(SetupInformation info)
        {
            var basePath = String.Format(@"{0}..\..\..\", AppDomain.CurrentDomain.BaseDirectory);

            await BuildAndDeployJob(info, "CheckAdminsJob",
                (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.CheckAdminsJob")).FullName,
                basePath, JobType.Triggered);
            await BuildAndDeployJob(info, "ExternalUsersJob",
                (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.ExternalUsersJob")).FullName,
                basePath, JobType.Triggered);
            await BuildAndDeployJob(info, "ScheduledJob",
                (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.ScheduledJob")).FullName,
                basePath, JobType.Triggered);
            await BuildAndDeployJob(info, "ContinousJob",
                (new System.IO.FileInfo(basePath + @"OfficeDevPnP.PartnerPack.ContinousJob")).FullName,
                basePath, JobType.Continuous);
        }

        private async static Task BuildAndDeployJob(SetupInformation info, String jobName, String jobPath, String basePath, JobType jobType)
        {
            // Run PowerShell script to build and deploy
            Hashtable packageBuildParameters = new Hashtable();
            packageBuildParameters.Add("ProjectPath", jobPath);

            var powerShellScriptPath = (new System.IO.FileInfo($@"{basePath}\OfficeDevPnP.PartnerPack.Setup\Scripts\MsBuildWebJob.ps1")).FullName;
            string packageBuildResult = Run.RunScript(powerShellScriptPath, packageBuildParameters);

            if (packageBuildResult.Contains("Build FAILED"))
            {
                // TODO: Handle exception
            }

            // Create the WebJob ZIP file
            var binPath = Path.Combine(jobPath, @"bin\Release");
            var zipPath = $@"{basePath}\OfficeDevPnP.PartnerPack.Setup\{jobName}.zip";

            // Remove any already existing file
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }

            System.IO.Compression.ZipFile.CreateFromDirectory(binPath, zipPath);

            // Get the Azure App Service Publishing Credentials
            var xmlPublishingSettings = XElement.Parse(info.AzureAppPublishingSettings);
            if (xmlPublishingSettings != null)
            {
                var xmlPublishProfile = xmlPublishingSettings.Element("publishProfile");

                if (xmlPublishProfile != null)
                {
                    var username = xmlPublishProfile.Attribute("userName").Value;
                    var password = xmlPublishProfile.Attribute("userPWD").Value;

                    // Upload the WebJobB
                    await AzureManagementUtility.UploadWebJob(info.AzureAppServiceName, username, password, jobName, zipPath, jobType);
                }
            }

            // Remove the ZIP file after having created the job
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
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
        CreateResourceGroup,
        CreateBlobStorageAccount,
        CreateAzureAppService,
        ConfigureSettings,
        ProvisionWebSite,
        ProvisionWebJobs,
        Completed,
    }

    public enum JobType
    {
        Triggered,
        Continuous,
    }
}
