using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Interface that defines the common behavior for any Sites Provisioning Templates Repository
    /// </summary>
    public interface ITemplateRepository : IConfigurable
    {
        /// <summary>
        /// Retrieves a single Provisioning Template
        /// </summary>
        /// <param name="templateUri">Defines the repository relative URI of the template in the target repository</param>
        /// <returns>Returns the Provisioning Template object for using it directly</returns>
        ProvisioningTemplate GetProvisioningTemplate(String templateUri);

        /// <summary>
        /// Allows to search for Provisioning Templates in the target repository
        /// </summary>
        /// <param name="searchText">Any free text to search for in title, abstract or URL/SEO of the template</param>
        /// <param name="platforms">The platforms to filter the provisioning templates</param>
        /// <param name="scope">The scope to filter the provisioning templates</param>
        /// <returns></returns>
        ProvisioningTemplateInformation[] SearchProvisioningTemplates(String searchText, TargetPlatform platforms, TargetScope scope);

        /// <summary>
        /// Saves a Provisioning Template into the target repository
        /// </summary>
        /// <param name="template">The Provisioning Template to save</param>
        /// <param name="templateUri">The repository relative URI of the provisioning template to save</param>
        void SaveProvisioningTemplate(ProvisioningTemplate template, String templateUri);

        /// <summary>
        /// Defines the display name for the current implementation of the ITemplateRepository
        /// </summary>
        String DisplayName { get;  }

        /// <summary>
        /// Defines whether the current templates repository can store new templates
        /// </summary>
        Boolean CanWrite { get; }
    }
}
