using Microsoft.Win32;
using OfficeDevPnP.PartnerPack.Setup.Components;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Security;
using System.Windows.Input;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Threading.Tasks;
using System.Configuration;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.PartnerPack.Setup.ViewModel
{

    public class MainViewModel : ValidableViewModelBase, IValidatableObject
    {
        private string _applicationName;
        private string _applicationUniqueUri;
        private string _applicationLogo;
        private string _azureWebAppUrl;
        private bool? _sslCertificateUpload;
        private string _sslCertificateFile;
        private string _sslCertificateCommonName;    
        private string _sslCertificatePassword;
        private DateTime _sslCertificateStartDate;
        private DateTime _sslCertificateEndDate;
        private string _absoluteUrl;
        private int _lcid;
        private int _timeZone;
        private string _primaryAdmin;
        private string _secondaryAdmin;
        private KeyValuePair<Guid, string>? _azureSubscription;
        private KeyValuePair<Guid, string>[] _azureSubscriptions;
        private KeyValuePair<string, string>? _azureLocation;
        private KeyValuePair<string, string>[] _azureLocations;
        private bool _azureLoggedIn;
        private string _azureAppServiceName;
        private string _azureBlobStorageName;
        private double _setupProgress;
        private bool _hasFatalError;
        private string _setupProgressDescription;
        private string _fatalErrorDescription;
        private bool _setupInProgress;
        private string _tenantName;

        private String _office365AccessToken;
        private Guid? _office365AzureSubscription;
        private String _azureAccessToken;

        public MainViewModel()
        {
            Office365LoginCommand = new ActionCommand(Office365Login);
            BrowseLogoCommand = new ActionCommand(BrowseLogo);
            BrowseCertificateCommand = new ActionCommand(BrowseCertificate);
            ResetLogoCommand = new ActionCommand(() => ApplicationLogo = null);
            AzureLoginCommand = new ActionCommand(AzureLogin);
            SetupCommand = new ActionCommand(Setup, () => !SetupInProgress);
            CloseFatalErrorCommand = new Components.ActionCommand(CloseFatalError, () => HasFatalError);
            SslCertificateUpload = true;
            SslCertificateStartDate = DateTime.Today;
            SslCertificateEndDate = DateTime.Today.AddYears(1);
            SslCertificateFile = "(Please select a file)";
            Lcid = 1033;
            TimeZone = 93;

            this.ApplicationName = "PnP Partner Pack";
            this.ApplicationLogo = new System.IO.FileInfo(
                $@"{AppDomain.CurrentDomain.BaseDirectory}..\..\..\..\Scripts\SharePoint-PnP-Icon.png").FullName;
            
            this.SslCertificateGenerate = true;
            this.SslCertificateFile = String.Empty;
            this.SslCertificateCommonName = "PnP-Partner-Pack";
            this.Lcid = 1033;
            this.TimeZone = 4;
        }

        private void CloseFatalError()
        {
            HasFatalError = false;
            FatalErrorDescription = "";
        }

        public ICommand Office365LoginCommand { get; }

        public ICommand BrowseLogoCommand { get; }

        public ICommand ResetLogoCommand { get; }

        public ICommand BrowseCertificateCommand { get; }

        public ICommand AzureLoginCommand { get; }

        public ActionCommand SetupCommand { get; }

        public ActionCommand CloseFatalErrorCommand { get; }

        public Guid? Office365AzureSubscription
        {
            get { return _office365AzureSubscription; }
            set
            {
                if (Set(ref _office365AzureSubscription, value))
                    ValidateModelProperty(value);
            }
        }

        [Required(ErrorMessage = "The application name is required")]
        public string ApplicationName
        {
            get { return _applicationName; }
            set
            {
                if (Set(ref _applicationName, value))
                    ValidateModelProperty(value);
            }
        }

        [Required(ErrorMessage = "The application unique URI is required")]
        [RegularExpression(@"https:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)", ErrorMessage = "URI must be formatted as https://site.com")]
        public string ApplicationUniqueUri
        {
            get { return _applicationUniqueUri; }
            set
            {
                if (Set(ref _applicationUniqueUri, value))
                    ValidateModelProperty(value);
            }
        }

        public string SslCertificateFile
        {
            get { return _sslCertificateFile; }
            set
            {
                if (Set(ref _sslCertificateFile, value))
                    ValidateModelProperty(value);
            }
        }

        public string ApplicationLogo
        {
            get { return _applicationLogo; }
            set
            {
                if (Set(ref _applicationLogo, value))
                    ValidateModelProperty(value);
            }
        }

        public bool SetupInProgress
        {
            get { return _setupInProgress; }
            private set
            {
                if (Set(ref _setupInProgress, value))
                {
                    SetupCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public bool? SslCertificateUpload
        {
            get { return _sslCertificateUpload; }
            set
            {
                if (Set(ref _sslCertificateUpload, value))
                {
                    SslCertificateFile = null;
                    SslCertificateCommonName = null;
                    OnPropertyChanged(nameof(SslCertificateGenerate));
                }
            }
        }

        public bool? SslCertificateGenerate
        {
            get { return !SslCertificateUpload; }
            set
            {
                SslCertificateUpload = !value;
            }
        }

        //[Required(ErrorMessage = "The Azure Web App URL is required")]
        //[RegularExpression(@"https:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)", ErrorMessage = "URI must be formatted as https://site.com")]
        public string AzureWebAppUrl
        {
            get { return _azureWebAppUrl; }
            set
            {
                if (Set(ref _azureWebAppUrl, value))
                    ValidateModelProperty(value);
            }
        }

        [Required(ErrorMessage = "The password is required")]
        [StringLength(100, MinimumLength = 6, ErrorMessage = "Password must be at least 6 characters")]
        public string SslCertificatePassword
        {
            get { return _sslCertificatePassword; }
            set
            {
                if (Set(ref _sslCertificatePassword, value))
                    ValidateModelProperty(value);
            }
        }

        public DateTime SslCertificateStartDate
        {
            get { return _sslCertificateStartDate; }
            set
            {
                if (Set(ref _sslCertificateStartDate, value))
                    ValidateModelProperty(value);
            }
        }

        public DateTime SslCertificateEndDate
        {
            get { return _sslCertificateEndDate; }
            set
            {
                if (Set(ref _sslCertificateEndDate, value))
                    ValidateModelProperty(value);
            }
        }

        [RequiredIf(nameof(SslCertificateGenerate), ErrorMessage = "The common name is required")]
        public string SslCertificateCommonName
        {
            get { return _sslCertificateCommonName; }
            set
            {
                if (Set(ref _sslCertificateCommonName, value))
                    ValidateModelProperty(value);
            }
        }

        [Required(ErrorMessage = "The absolute URL is required")]
        [RegularExpression(@"https:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)", ErrorMessage = "URL must be formatted as https://site.com")]
        public string AbsoluteUrl
        {
            get { return _absoluteUrl; }
            set
            {
                if (Set(ref _absoluteUrl, value))
                    ValidateModelProperty(value);
            }
        }

        public int Lcid
        {
            get { return _lcid; }
            set
            {
                if (Set(ref _lcid, value))
                    ValidateModelProperty(value);
            }
        }

        public int TimeZone
        {
            get { return _timeZone; }
            set
            {
                if (Set(ref _timeZone, value))
                    ValidateModelProperty(value);
            }
        }

        public string PrimaryAdmin
        {
            get { return _primaryAdmin; }
            set
            {
                if (Set(ref _primaryAdmin, value))
                    ValidateModelProperty(value);
            }
        }

        public string SecondaryAdmin
        {
            get { return _secondaryAdmin; }
            set
            {
                if (Set(ref _secondaryAdmin, value))
                    ValidateModelProperty(value);
            }
        }

        public KeyValuePair<Guid, string>[] AzureSubscriptions
        {
            get { return _azureSubscriptions; }
            private set
            {
                if (Set(ref _azureSubscriptions, value))
                {
                    ValidateModelProperty(value);
                    OnPropertyChanged(nameof(AzureSubscriptionsReady));
                    OnPropertyChanged(nameof(AzureSubscriptionsNotReady));
                }
            }
        }

        [Required(ErrorMessage = "The Azure subscription is required")]
        public KeyValuePair<Guid, string>? AzureSubscription
        {
            get { return _azureSubscription; }
            set
            {
                if (Set(ref _azureSubscription, value))
                    ValidateModelProperty(value);
            }
        }

        public KeyValuePair<string, string>[] AzureLocations
        {
            get { return _azureLocations; }
            private set
            {
                if (Set(ref _azureLocations, value))
                {
                    ValidateModelProperty(value);
                    OnPropertyChanged(nameof(AzureLocationsReady));
                }
            }
        }

        public KeyValuePair<string, string>? AzureLocation
        {
            get { return _azureLocation; }
            set
            {
                if (Set(ref _azureLocation, value))
                    ValidateModelProperty(value);
            }
        }

        public double SetupProgress
        {
            get { return _setupProgress; }
            set
            {
                if (Set(ref _setupProgress, value))
                    ValidateModelProperty(value);
            }
        }

        public bool HasFatalError
        {
            get { return _hasFatalError; }
            set
            {
                if (Set(ref _hasFatalError, value))
                    CloseFatalErrorCommand.RaiseCanExecuteChanged();
            }
        }

        public string FatalErrorDescription
        {
            get { return _fatalErrorDescription; }
            set
            {
                Set(ref _fatalErrorDescription, value);
            }
        }

        public string SetupProgressDescription
        {
            get { return _setupProgressDescription; }
            set
            {
                Set(ref _setupProgressDescription, value);
            }
        }

        public bool AzureSubscriptionsReady => _azureLoggedIn && _azureSubscriptions != null;

        public bool AzureSubscriptionsNotReady => _azureLoggedIn && _azureSubscriptions == null;

        public bool AzureLocationsReady => _azureLocations != null && _azureLocations.Length > 0;

        [Required(ErrorMessage = "The Azure App Service name is required")]
        [RegularExpression(@"^[-a-z][-a-z0-9]{2,62}$", ErrorMessage = "Name must be lower case and must contain letters, numbers, and dashes only")]
        public string AzureAppServiceName
        {
            get { return _azureAppServiceName; }
            set
            {
                if (Set(ref _azureAppServiceName, value))
                    ValidateModelProperty(value);
            }
        }

        [Required(ErrorMessage = "The Azure Blob Storage name is required")]
        [RegularExpression(@"^[a-z][a-z0-9]{2,62}$", ErrorMessage = "Name must be lower case and must contain letters and numbers only")]
        public string AzureBlobStorageName
        {
            get { return _azureBlobStorageName; }
            set
            {
                if (Set(ref _azureBlobStorageName, value))
                    ValidateModelProperty(value);
            }
        }

        private void BrowseCertificate()
        {
            var dialog = new OpenFileDialog();
            dialog.Title = "Select the certificate to use";
            dialog.Filter = "Certificate|*.pfx";
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                SslCertificateFile = dialog.FileName;
            }
        }

        private void BrowseLogo()
        {
            var dialog = new OpenFileDialog();
            dialog.Title = "Select a logo for your application";
            dialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;";
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                ApplicationLogo = dialog.FileName;
            }
        }

        private async void Setup()
        {
            ValidateModel();
            if (HasErrors)
            {
                return;
            }

            SetupInProgress = true;

            try
            {
                try
                {
                    await Task.Run(() =>
                    {
                        SetupManager.SetupPartnerPackAsync(MapSetupInformation(this)).Wait();
                    });

                    // Start the browser targeting the PnP Partner Pack site
                    System.Diagnostics.Process.Start(this.AzureWebAppUrl);
                }
                finally
                {
                    SetupInProgress = false;
                }
            }
            catch (Exception ex)
            {
                HasFatalError = true;
                FatalErrorDescription = ((ex as AggregateException)?.InnerException ?? ex).ToString();
            }
        }

        private SetupInformation MapSetupInformation(MainViewModel viewModel)
        {
            var result = new SetupInformation
            {
                ViewModel = viewModel,
                AzureAccessToken = viewModel._azureAccessToken,
                Office365AccessToken = viewModel._office365AccessToken,
                Office365TargetSubscriptionId = viewModel.Office365AzureSubscription.HasValue ?
                        viewModel.Office365AzureSubscription.Value : Guid.Empty,
                ApplicationName = viewModel.ApplicationName,
                ApplicationUniqueUri = viewModel.ApplicationUniqueUri,
                ApplicationLogoPath = viewModel.ApplicationLogo,
                AzureWebAppUrl = $"https://{viewModel.AzureAppServiceName}.azurewebsites.net/",
                SslCertificateGenerate = viewModel.SslCertificateGenerate.HasValue ?
                        viewModel.SslCertificateGenerate.Value : false,
                SslCertificateFile = viewModel.SslCertificateFile,
                SslCertificatePassword = viewModel.SslCertificatePassword,
                SslCertificateCommonName = viewModel.SslCertificateCommonName,
                SslCertificateStartDate = viewModel.SslCertificateStartDate,
                SslCertificateEndDate = viewModel.SslCertificateEndDate,
                InfrastructuralSiteUrl = viewModel.AbsoluteUrl,
                InfrastructuralSiteLCID = viewModel.Lcid,
                InfrastructuralSiteTimeZone = viewModel.TimeZone,
                InfrastructuralSitePrimaryAdmin = viewModel.PrimaryAdmin,
                InfrastructuralSiteSecondaryAdmin = viewModel.SecondaryAdmin,
                AzureTargetSubscriptionId = viewModel.AzureSubscription.HasValue ?
                        viewModel.AzureSubscription.Value.Key : Guid.Empty,
                AzureLocationId = viewModel.AzureLocation.HasValue ?
                        viewModel.AzureLocation.Value.Key : String.Empty,
                AzureLocationDisplayName = viewModel.AzureLocation.HasValue ?
                        viewModel.AzureLocation.Value.Value : String.Empty,
                AzureAppServiceName = viewModel.AzureAppServiceName,
                AzureBlobStorageName = viewModel.AzureBlobStorageName,
                AzureADTenant = viewModel._tenantName,
            };

            this.AzureWebAppUrl = result.AzureWebAppUrl;

            return (result);
        }

        private async void Office365Login()
        {
            // Get the list of subscriptions for the current user
            _office365AccessToken = await AzureManagementUtility.GetAccessTokenAsync(
                AzureManagementUtility.AzureManagementApiURI, 
                ConfigurationManager.AppSettings["O365:ClientId"]);

            try
            {
                var office365Account = await AzureManagementUtility.GetUserUniqueId(ConfigurationManager.AppSettings["O365:ClientId"]);

                Regex regex = new Regex(@"(?<mailbox>.*)@(?<tenant>\w*)\.(?<remainder>.*)");
                var match = regex.Match(office365Account);
                _tenantName = match.Groups["tenant"].Value;
                var remainderName = match.Groups["remainder"].Value;

                // Configure parameters based on the current Office 365 tenant name
                this.ApplicationUniqueUri = $"https://{_tenantName}.{remainderName}/PnP-Partner-Pack";
                // this.AzureWebAppUrl = $"https://pnp-partner-pack-{_tenantName}.azurewebsites.net/";
                this.AbsoluteUrl = $"https://{_tenantName}.sharepoint.com/sites/PnP-Partner-Pack";
                this.PrimaryAdmin = office365Account;
                this.AzureAppServiceName = $"pnp-partner-pack-{_tenantName}";
                this.AzureBlobStorageName = $"{_tenantName}storage";
            }
            catch
            {
                // Intentionally ignore any exception related to settings suggestion, because they are not really critical
            }

            if (!String.IsNullOrEmpty(_office365AccessToken))
            {
                var subscriptions = (await AzureManagementUtility.ListSubscriptionsAsync(_office365AccessToken)).ToArray();
                if (subscriptions.Length == 0)
                {
                    throw new ApplicationException("Missing default Azure subscription for Office 365 tenant!");
                }
                else
                {
                    Office365AzureSubscription = subscriptions[0].Key;
                }
            }
        }

        private async void AzureLogin()
        {
            // Start the ring rotation
            _azureLoggedIn = true;
            AzureSubscriptions = null;
            OnPropertyChanged(nameof(AzureSubscriptionsReady));
            OnPropertyChanged(nameof(AzureSubscriptionsNotReady));

            // Get the list of subscriptions for the current user
            //AzureSubscriptions = Enumerable.Range(1, 10).Select(n => new KeyValuePair<Guid, string>(Guid.NewGuid(), "Subscription " + n)).ToArray();
            _azureAccessToken = await AzureManagementUtility.GetAccessTokenAsync(
                AzureManagementUtility.AzureManagementApiURI);
            if (!String.IsNullOrEmpty(_azureAccessToken))
            {
                AzureSubscriptions = (await AzureManagementUtility.ListSubscriptionsAsync(_azureAccessToken)).ToArray();
                AzureSubscription = AzureSubscriptions[0];

                AzureLocations = (await AzureManagementUtility.ListLocations(_azureAccessToken, AzureSubscription.Value.Key)).ToArray();
                AzureLocation = AzureLocations[0];
            }
        }

        IEnumerable<ValidationResult> IValidatableObject.Validate(ValidationContext validationContext)
        {
            if (SslCertificateUpload.GetValueOrDefault() && String.IsNullOrWhiteSpace(SslCertificateFile))
                yield return new ValidationResult("Please select a certificate to use", new[] { nameof(SslCertificateFile) });
        }
    }
}