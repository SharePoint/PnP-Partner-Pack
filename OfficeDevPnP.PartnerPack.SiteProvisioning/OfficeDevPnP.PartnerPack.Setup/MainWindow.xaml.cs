using OfficeDevPnP.PartnerPack.Setup.Components;
using OfficeDevPnP.PartnerPack.Setup.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OfficeDevPnP.PartnerPack.Setup
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private MainViewModel ViewModel => ((MainViewModel)DataContext);

        private void Setup_Click(object sender, RoutedEventArgs e)
        {
            // Find any element with errors
            var elementWithErrors = GetDescendants(this).FirstOrDefault(o => Validation.GetErrors(o).Count > 0) as FrameworkElement;
            if (elementWithErrors != null)
            {
                // Move scroll viewer to the element
                var t = elementWithErrors.TransformToAncestor(scrollViewer);
                scrollViewer.ScrollToVerticalOffset(t.Transform(new System.Windows.Point(0, -30)).Y);

                // Move focus on the element
                Keyboard.Focus(elementWithErrors);
            }

            if (!SetupManager.SetupPartnerPack(new SetupInformation {
                ApplicationName = this.ViewModel.ApplicationName,
                ApplicationUniqueUri = this.ViewModel.ApplicationUniqueUri,
                ApplicationLogoPath = this.ViewModel.ApplicationLogo,
                AzureWebAppUrl = this.ViewModel.AzureWebAppUrl,
                SslCertificateGenerate = this.ViewModel.SslCertificateGenerate.HasValue ? 
                    this.ViewModel.SslCertificateGenerate.Value : false,
                SslCertificateFile = this.ViewModel.SslCertificateFile,
                SslCertificatePassword = this.ViewModel.SslCertificatePassword,
                SslCertificateCommonName = this.ViewModel.SslCertificateCommonName,
                SslCertificateStartDate = this.ViewModel.SslCertificateStartDate,
                SslCertificateEndDate = this.ViewModel.SslCertificateEndDate,
                InfrastructuralSiteUrl = this.ViewModel.AbsoluteUrl,
                InfrastructuralSiteLCID = this.ViewModel.Lcid,
                InfrastructuralSiteTimeZone = this.ViewModel.TimeZone,
                AzureTargetSubscription = this.ViewModel.AzureSubscription.HasValue ? 
                    this.ViewModel.AzureSubscription.Value.Key : Guid.Empty,
                AzureAppServiceName = this.ViewModel.AzureAppServiceName,
                AzureBlobStorageName = this.ViewModel.AzureBlobStorageName,
            }))
            {
                // Show any error on the screen/log ...

            }
        }

        private IEnumerable<DependencyObject> GetDescendants(DependencyObject d)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(d); i++)
            {
                var c = VisualTreeHelper.GetChild(d, i);
                yield return c;

                foreach (var dc in GetDescendants(c))
                    yield return dc;
            }
        }
    }
}
