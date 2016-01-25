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

                StoreName storeName;
                StoreLocation storeLocation;

                Enum.TryParse(_configuration.CertificateSettings.storeName,
                    out storeName);
                Enum.TryParse(_configuration.CertificateSettings.storeLocation,
                    out storeLocation);

                X509Store certStore = new X509Store(storeName, storeLocation);
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

        private static IReadOnlyDictionary<Type, ProvisioningJobHandler> GetJobHandlersByExecutionModel(
            PnPPartnerPackConfigurationProvisioningJobsJobTypeExecutionModel executionModel)
        {
            Dictionary<Type, ProvisioningJobHandler> handlers = new Dictionary<Type, ProvisioningJobHandler>();

            // Browse through the configured Job Handlers
            if (_configuration.ProvisioningJobs != null &&
                _configuration.ProvisioningJobs.JobHandlers != null)
            {
                foreach (var jobHandlerKind in _configuration.ProvisioningJobs.JobHandlers)
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

                    if (_configuration.ProvisioningJobs != null &&
                        _configuration.ProvisioningJobs.JobTypes != null)
                    {
                        // For each Job type associated with the current Job Handler
                        foreach (var jobKind in _configuration.ProvisioningJobs.JobTypes
                            .Where(j => j.executionModel == executionModel &&  j.handler == jobHandlerKind.name))
                        {
                            // Associate the Job type with the Job Handler
                            Type jobType = Type.GetType(jobKind.type, true);
                            handlers.Add(jobType, jobHandler);
                        }
                    }
                }
            }

            return (handlers);
        }

        private static readonly Lazy<IReadOnlyDictionary<Type, ProvisioningJobHandler>> _scheduledJobHandlers =
            new Lazy<IReadOnlyDictionary<Type, ProvisioningJobHandler>>(() => 
            GetJobHandlersByExecutionModel(PnPPartnerPackConfigurationProvisioningJobsJobTypeExecutionModel.Scheduled));

        private static readonly Lazy<IReadOnlyDictionary<Type, ProvisioningJobHandler>> _continousJobHandlers =
            new Lazy<IReadOnlyDictionary<Type, ProvisioningJobHandler>>(() =>
            GetJobHandlersByExecutionModel(PnPPartnerPackConfigurationProvisioningJobsJobTypeExecutionModel.Continous));

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

        /// <summary>
        /// Provides the list of Job Handlers running Continously
        /// </summary>
        public static IReadOnlyDictionary<Type, ProvisioningJobHandler> ContinousJobHandlers
        {
            get
            {
                return (_continousJobHandlers.Value);
            }
        }

        /// <summary>
        /// Provides the list of Job Handlers running based on a Schedule
        /// </summary>
        public static IReadOnlyDictionary<Type, ProvisioningJobHandler> ScheduledJobHandlers
        {
            get
            {
                return (_scheduledJobHandlers.Value);
            }
        }

        /// <summary>
        /// The URL of the application logo in the UI.
        /// </summary>
        /// <remarks>
        /// It is an optional attribute and by default the PnP logo will be used.
        /// </remarks>
        public static String LogoUrl
        {
            get
            {
                if (!String.IsNullOrEmpty(_configuration.GeneralSettings.LogoUrl))
                {
                    return (_configuration.GeneralSettings.LogoUrl);
                }
                else
                {
                    return ("/AppIcon.png");
                }
            }
        }

        /// <summary>
        /// The Title of the application in the UI.
        /// </summary>
        /// <remarks>
        /// It is an optional attribute and by default the name "PnP Partner Pack" will be used.
        /// </remarks>
        public static String Title
        {
            get
            {
                if (!String.IsNullOrEmpty(_configuration.GeneralSettings.Title))
                {
                    return (_configuration.GeneralSettings.Title);
                }
                else
                {
                    return ("PnP Partner Pack");
                }
            }
        }

        /// <summary>
        /// The Welcome Message used in the Home Page of the application. 
        /// </summary>
        public static String WelcomeMessage
        {
            get
            {
                if (_configuration.GeneralSettings.WelcomeMessage != null)
                {
                    return (_configuration.GeneralSettings.WelcomeMessage.InnerText);
                }
                else
                {
                    return (String.Empty);
                }
            }
        }

        /// <summary>
        /// The Footer Message used in the pages of the application.
        /// </summary>
        public static String FooterMessage
        {
            get
            {
                if (_configuration.GeneralSettings.FooterMessage != null)
                {
                    return (_configuration.GeneralSettings.FooterMessage.InnerText);
                }
                else
                {
                    return (String.Empty);
                }
            }
        }

        /// <summary>
        /// The name of the Site Template to use while creating new Sites
        /// </summary>
        public static String DefaultSiteTemplate
        {
            get
            {
                return (_configuration.GeneralSettings.defaultSiteTemplate);
            }
        }

        /// <summary>
        /// The name of the Azure Storage Queue used for Continously running Jobs
        /// </summary>
        public static String StorageQueueConnectionString
        {
            get
            {
                if (ConfigurationManager.ConnectionStrings["AzureWebJobsStorage"] != null)
                {
                    return (ConfigurationManager.ConnectionStrings["AzureWebJobsStorage"].ConnectionString);
                }
                else
                {
                    return (null);
                }
            }
        }

        /// <summary>
        /// The name of the Azure Storage Queue used for Continously running Jobs
        /// </summary>
        public const String StorageQueueName = "pnppartnerpackjobsqueue";
    }
}