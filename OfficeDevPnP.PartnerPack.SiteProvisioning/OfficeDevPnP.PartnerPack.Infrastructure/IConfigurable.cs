using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Defines the behavior for a configurable object instance
    /// </summary>
    public interface IConfigurable
    {
        /// <summary>
        /// Initializes the Configurable object instance
        /// </summary>
        void Init(XElement configuration);
    }
}
