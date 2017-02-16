using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows;

namespace OfficeDevPnP.PartnerPack.Setup
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            bool isAdmin;

            try
            {
                using (var user = WindowsIdentity.GetCurrent())
                {
                    WindowsPrincipal principal = new WindowsPrincipal(user);
                    isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
                }
            }
            catch (Exception)
            {
                isAdmin = false;
            }

            if (!isAdmin)
            {
                MessageBox.Show("This application requires administrative privileges to be executed!");
                this.Shutdown();
            }
        }
    }
}
