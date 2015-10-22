using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Configuration
{
    public class PnPPartnerPackConfigurationSectionHandler : IConfigurationSectionHandler
    {
        public object Create(object parent, object configContext, XmlNode section)
        {
            XmlNodeReader xnr = new XmlNodeReader(section);
            XmlSerializer xs = new XmlSerializer(typeof(PnPPartnerPackConfiguration));
            return (xs.Deserialize(xnr));
        }
    }
}
