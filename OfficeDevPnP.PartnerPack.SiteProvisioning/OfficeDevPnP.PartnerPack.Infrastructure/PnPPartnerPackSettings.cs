using OfficeDevPnP.PartnerPack.Infrastructure.Configuration;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Web;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System.Xml;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    public static class PnPPartnerPackSettings
    {
        private static String _clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static String _clientSecret = ConfigurationManager.AppSettings["ida:ClientSecret"];
        private static String _aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];

        private static PnPPartnerPackConfiguration _configuration =
            (PnPPartnerPackConfiguration)ConfigurationManager.GetSection("PnPPartnerPackConfiguration");

        private static String _tenant = _configuration.TenantSettings.tenant;
        private static String _infrastructureSiteUrl = _configuration.TenantSettings.infrastructureSiteUrl;
        private static String _provisioningRepositoryType = _configuration.ProvisioningRepository.type;

        private static readonly Lazy<X509Certificate2> _appOnlyCertificateLazy =
            new Lazy<X509Certificate2>(() => {

                X509Certificate2 appOnlyCertificate = null;

                X509Store certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                certStore.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certCollection = certStore.Certificates.Find(
                    X509FindType.FindByThumbprint,
                    _configuration.TenantSettings.appOnlyCertificateThumbprint,
                    false);

                // Get the first cert with the thumbprint
                if (certCollection.Count > 0)
                {
                    appOnlyCertificate = certCollection[0];
                }
                certStore.Close();

                return (appOnlyCertificate);
            });

        private static readonly Lazy<IReadOnlyDictionary<Type, ProvisioningJobHandler>> _jobHandlers =
            new Lazy<IReadOnlyDictionary<Type, ProvisioningJobHandler>>(() => {

                Dictionary<Type, ProvisioningJobHandler> handlers = new Dictionary<Type, ProvisioningJobHandler>();

                // Browse through the configured Job Handlers
                foreach (var jobHandlerKind in _configuration.JobsHandlers)
                {
                    // Create an instance of the Job Handler
                    Type jobHandlerType = Type.GetType(jobHandlerKind.type, true);
                    ProvisioningJobHandler jobHandler = (ProvisioningJobHandler)Activator.CreateInstance(jobHandlerType);

                    // If there is any configuration XML element
                    if (jobHandlerKind.Configuration != null)
                    {
                        // Convert it into a XElement
                        using (XmlReader reader = new XmlNodeReader(jobHandlerKind.Configuration))
                        {
                            XElement configuration = XElement.Load(reader);
 
                            // Initialize the Job Handler
                            jobHandler.Init(configuration);
                        }
                    }

                    if (jobHandlerKind.Jobs != null)
                    {
                        // For each Job type associated with the current Job Handler
                        foreach (var jobKind in jobHandlerKind.Jobs)
                        {
                            // Associate the Job type with the Job Handler
                            Type jobType = Type.GetType(jobKind.type, true);
                            handlers.Add(jobType, jobHandler);
                        }
                    }
                }

                return (handlers);
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
        /// Provides the the target Tenant for the PnP Partner Pack
        /// </summary>
        public static String Tenant
        {
            get
            {
                return (_tenant);
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
        /// Provides the XElement configuration for the Provisioning Repository, if any
        /// </summary>
        public static XElement ProvisioningRepositoryConfiguration
        {
            get
            {
                if (_configuration.ProvisioningRepository.Configuration != null)
                {
                    using (XmlReader reader = new XmlNodeReader(_configuration.ProvisioningRepository.Configuration))
                    {
                        XElement result = XElement.Load(reader);
                        return (result);
                    }
                }
                return (null);
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

        public static IReadOnlyDictionary<Type, ProvisioningJobHandler> JobHandlers
        {
            get
            {
                return (_jobHandlers.Value);
            }
        }
    }
}