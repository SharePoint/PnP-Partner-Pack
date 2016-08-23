using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    /// <summary>
    /// Provides the basic interface for any repository of templates
    /// </summary>
    public interface ITemplatesProvider : IConfigurable
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
        /// Defines the display name for the current implementation of the ITemplateRepository
        /// </summary>
        String DisplayName { get; }
    }
}
