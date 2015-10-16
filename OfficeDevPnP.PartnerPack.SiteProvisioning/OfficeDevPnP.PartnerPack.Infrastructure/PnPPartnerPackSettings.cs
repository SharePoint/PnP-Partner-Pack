using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Web;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    public static class PnPPartnerPackSettings
    {
        private static String _clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static String _clientSecret = ConfigurationManager.AppSettings["ida:ClientSecret"];
        private static String _aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private static String _infrastructureSiteUrl = ConfigurationManager.AppSettings["pnp:InfrastructureSiteUrl"];
        private static String _provisioningRepositoryType = ConfigurationManager.AppSettings["pnp:ProvisioningRepositoryType"];

        private static readonly Lazy<X509Certificate2> _appOnlyCertificateLazy =
            new Lazy<X509Certificate2>(() => {

                X509Certificate2 appOnlyCertificate = null;

                X509Store certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                certStore.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certCollection = certStore.Certificates.Find(
                    X509FindType.FindByThumbprint,
                    ConfigurationManager.AppSettings["pnp:AppOnlyCertificateThumbprint"],
                    false);

                // Get the first cert with the thumbprint
                if (certCollection.Count > 0)
                {
                    appOnlyCertificate = certCollection[0];
                }
                certStore.Close();

                return (appOnlyCertificate);
            });

        /// <summary>
        /// Provides the Azure AD Client ID
        /// </summary>
        public static String ClientId
        {
            get {
                return (_clientId);
            }
        }

        /// <summary>
        /// Provides the Azure AD Client Secret
        /// </summary>
        public static String ClientSecret
        {
            get
            {
                return (_clientSecret);
            }
        }

        /// <summary>
        /// Provides the Azure AD Instance URL
        /// </summary>
        public static String AADInstance
        {
            get
            {
                return (_aadInstance);
            }
        }

        /// <summary>
        /// Provides the URL of the PnP Partner Pack Infrastructural Site
        /// </summary>
        public static String InfrastructureSiteUrl
        {
            get
            {
                return (_infrastructureSiteUrl);
            }
        }
        
        /// <summary>
        /// Provides the .NET type name of the Provisioning Repository
        /// </summary>
        public static String ProvisioningRepositoryType
        {
            get
            {
                return (_provisioningRepositoryType);
            }
        }

        /// <summary>
        /// Provides the X.509 certificate for Azure AD AppOnly Authentication
        /// </summary>
        public static X509Certificate2 AppOnlyCertificate
        {
            get
            {
                return (_appOnlyCertificateLazy.Value);
            }
        }
    }
}