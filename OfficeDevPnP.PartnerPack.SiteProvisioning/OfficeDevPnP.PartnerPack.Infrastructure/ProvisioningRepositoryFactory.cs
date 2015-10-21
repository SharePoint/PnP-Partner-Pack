using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Provides the factory to get the current ProvisioningRepository
    /// </summary>
    public static class ProvisioningRepositoryFactory
    {
        private static readonly Lazy<IProvisioningRepository> _provisioningRepositoryLazy =
            new Lazy<IProvisioningRepository>(() => {

                // Creates the Provisioning Repository instance
                Type provisioningRepositoryType = Type.GetType(PnPPartnerPackSettings.ProvisioningRepositoryType, true);
                IProvisioningRepository result = (IProvisioningRepository)Activator.CreateInstance(provisioningRepositoryType);

                // Initializes the Provisioning Repository instance
                result.Init();

                return (result);
            });

        /// <summary>
        /// Provides a singleton reference to the current Provisioning Repository
        /// </summary>
        public static IProvisioningRepository Current {
            get { return (_provisioningRepositoryLazy.Value); }
        }
    }
}
