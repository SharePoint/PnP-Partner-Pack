using OfficeDevPnP.PartnerPack.Setup.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    /// <summary>
    /// Defines the setup settings
    /// </summary>
    public class SetupInformation
    {
        public MainViewModel ViewModel { get; set; }
        public Guid Office365TargetSubscriptionId { get; set; }
        public String ApplicationName { get; set; }
        public String ApplicationUniqueUri { get; set; }
        public String ApplicationLogoPath { get; set; }
        public String AzureWebAppUrl { get; set; }
        public Boolean SslCertificateGenerate { get; set; }
        public String SslCertificateFile { get; set; }
        public String SslCertificatePassword { get; set; }
        public String SslCertificateCommonName { get; set; }
        public DateTime SslCertificateStartDate { get; set; }
        public DateTime SslCertificateEndDate { get; set; }
        public String SslCertificateThumbprint { get; set; }
        public String InfrastructuralSiteUrl { get; set; }
        public Int32 InfrastructuralSiteLCID { get; set; }
        public Int32 InfrastructuralSiteTimeZone { get; set; }
        public String InfrastructuralSitePrimaryAdmin { get; set; }
        public String InfrastructuralSiteSecondaryAdmin { get; set; }
        public Guid AzureTargetSubscriptionId { get; set; }
        public String AzureLocationId { get; set; }
        public String AzureLocationDisplayName { get; set; }
        public String AzureAppServiceName { get; set; }
        public String AzureBlobStorageName { get; set; }
        public String Office365AccessToken { get; set; }
        public String AzureAccessToken { get; set; }
        public Guid AzureAppClientId { get; set; }
        public String AzureAppSharedSecret { get; set; }
        public String AzureAppKeyCredential { get; set; }
        public String AzureResourceGroupName { get; set; }
        public String AzureServicePlanName { get; set; }
        public String AzureStorageAccountName { get; set; }
        public String AzureStorageKey { get; set; }
        public X509Certificate2 AuthenticationCertificate { get; set; }
        public String AzureAppPublishingSettings { get; set; }
        public String AzureADTenant { get; set; }
    }
}
