using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class Key
    {
        public string keyName { get; set; }
        public string permissions { get; set; }
        public string value { get; set; }
    }

    public class StorageKeys
    {
        public List<Key> keys { get; set; }
    }
}
