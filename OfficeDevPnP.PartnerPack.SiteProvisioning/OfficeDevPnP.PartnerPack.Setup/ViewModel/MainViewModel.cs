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
        private KeyValuePair<Guid, string>? _azureSubscription;
        private KeyValuePair<Guid, string>[] _azureSubscriptions;
        private bool _azureLoggedIn;
        private string _azureAppServiceName;
        private string _azureBlobStorageName;
        private double _setupProgress;
        private string _setupProgressDescription;
        private bool _setupInProgress;

        public MainViewModel()
        {
            BrowseLogoCommand = new ActionCommand(BrowseLogo);
            BrowseCertificateCommand = new ActionCommand(BrowseCertificate);
            ResetLogoCommand = new ActionCommand(() => ApplicationLogo = null);
            AzureLoginCommand = new ActionCommand(AzureLogin);
            SetupCommand = new ActionCommand(Setup, () => !SetupInProgress);
            SslCertificateUpload = true;
            SslCertificateStartDate = DateTime.Today;
            SslCertificateEndDate = DateTime.Today.AddYears(1);
            SslCertificateFile = "(Please select a file)";
            Lcid = 1033;
            TimeZone = 93;
        }

        public ICommand BrowseLogoCommand { get; }

        public ICommand ResetLogoCommand { get; }

        public ICommand BrowseCertificateCommand { get; }

        public ICommand AzureLoginCommand { get; }

        public ICommand SetupCommand { get; }

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
        [RegularExpression(@"https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)", ErrorMessage = "URI must be formatted as http://site.com")]
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
                    OnPropertyChanged(nameof(SetupCommand));
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

        [Required(ErrorMessage = "The Azure Web App URL is required")]
        [RegularExpression(@"https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)", ErrorMessage = "URI must be formatted as https://site.com")]
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
        [RegularExpression(@"https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)", ErrorMessage = "URL must be formatted as http://site.com")]
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

        public double SetupProgress
        {
            get { return _setupProgress; }
            private set
            {
                Set(ref _setupProgress, value);
            }
        }

        public string SetupProgressDescription
        {
            get { return _setupProgressDescription; }
            private set
            {
                Set(ref _setupProgressDescription, value);
            }
        }

        public bool AzureSubscriptionsReady => _azureLoggedIn && _azureSubscriptions != null;

        public bool AzureSubscriptionsNotReady => _azureLoggedIn && _azureSubscriptions == null;

        [Required(ErrorMessage = "The Azure App Service name is required")]
        [RegularExpression(@"^[a-z][a-z0-9]{2,62}$", ErrorMessage = "Name must be lower case and must contain letters and numbers only")]
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
                // TODO: setup
                for (int i = 0; i < 100; i++)
                {
                    SetupProgress = i;
                    SetupProgressDescription = $"Progress {i}%";

                    await Task.Delay(100);
                }

                MessageBox.Show("Done!");
            }
            finally
            {
                SetupInProgress = false;
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
            var accessToken = await AzureManagementUtility.GetAccessTokenAsync();
            AzureSubscriptions = (await AzureManagementUtility.ListSubscriptionsAsync(accessToken)).ToArray();
            AzureSubscription = AzureSubscriptions[0];
        }

        IEnumerable<ValidationResult> IValidatableObject.Validate(ValidationContext validationContext)
        {
            if (SslCertificateUpload.GetValueOrDefault() && String.IsNullOrWhiteSpace(SslCertificateFile))
                yield return new ValidationResult("Please select a certificate to use", new[] { nameof(SslCertificateFile) });
        }
    }
}