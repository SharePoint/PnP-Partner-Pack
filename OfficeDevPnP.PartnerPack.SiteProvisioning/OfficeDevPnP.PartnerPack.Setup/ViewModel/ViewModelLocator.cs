using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.ViewModel
{
    public class ViewModelLocator
    {
        private static MainViewModel _main = new MainViewModel();

        public MainViewModel Main => _main;
    }
}
