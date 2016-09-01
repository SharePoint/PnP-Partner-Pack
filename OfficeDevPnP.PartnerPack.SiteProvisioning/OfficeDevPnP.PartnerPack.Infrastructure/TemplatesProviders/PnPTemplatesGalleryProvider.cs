using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml.Linq;

namespace OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders
{
    /// <summary>
    /// Implements the Templates Provider that uses the OfficeDev PnP Templates Gallery as source
    /// </summary>
    public class PnPTemplatesGalleryProvider : ITemplatesProvider
    {
        private String _templatesGalleryBaseUrl;

        public string DisplayName
        {
            get { return ("PnP Templates Gallery"); }
        }

        public void Init(XElement configuration)
        {
            // TODO: read from XElement configuration the base URL of the Templates Gallery site
            // Mind the final / in the provided URL
            // <gallery url="https://templates-gallery.officedevpnp.com/" />

            if (configuration.Name != "{http://schemas.dev.office.com/PnP/2016/08/PnPPartnerPackConfiguration}gallery")
            {
                throw new ApplicationException("Invalid configuration settings for PnPTemplatesGalleryProvider, missing gallery root element!");
            }

            var urlAttribute = configuration.Attribute("url");
            if (urlAttribute == null)
            {
                throw new ApplicationException("Invalid configuration settings for PnPTemplatesGalleryProvider, missing url attribute!");
            }

            this._templatesGalleryBaseUrl = urlAttribute.Value;

            if (!this._templatesGalleryBaseUrl.EndsWith("/"))
            {
                this._templatesGalleryBaseUrl += "/";
            }
        }

        public ProvisioningTemplate GetProvisioningTemplate(string templateUri)
        {
            ProvisioningTemplate result = null;

            // Get the template via HTTP REST
            var templateStream = HttpHelper.MakeGetRequestForStream(
                $"{this._templatesGalleryBaseUrl}api/DownloadTemplate?templateUri={HttpUtility.UrlEncode(templateUri)}",
                "application/octet-stream");

            // If we have any result
            if (templateStream != null)
            {
                XMLTemplateProvider provider = new XMLOpenXMLTemplateProvider(
                    new OpenXMLConnector(templateStream));

                var openXMLFileName = templateUri.Substring(templateUri.LastIndexOf("/") + 1);

                // Determine the name of the XML file inside the PNP Open XML file
                var xmlTemplateFile = openXMLFileName.ToLower().Replace(".pnp", ".xml");

                // Get the template
                result = provider.GetTemplate(xmlTemplateFile);
                result.Connector = provider.Connector;
            }

            return (result);
        }

        /// <summary>
        /// Search for templates in the PnP Templates Gallery
        /// </summary>
        /// <param name="searchText"></param>
        /// <param name="platforms"></param>
        /// <param name="scope"></param>
        /// <returns></returns>
        public ProvisioningTemplateInformation[] SearchProvisioningTemplates(string searchText, TargetPlatform platforms, TargetScope scope)
        {
            ProvisioningTemplateInformation[] result = null;

            String targetPlatforms = platforms.ToString();
            String targetScopes = scope.ToString();

            // Search via HTTP REST
            var jsonSearchResult = HttpHelper.MakeGetRequestForString(
                $"{this._templatesGalleryBaseUrl}api/SearchTemplates?searchText={HttpUtility.UrlEncode(searchText)}&platforms={targetPlatforms}&scope={targetScopes}");

            // If we have any result
            if (!String.IsNullOrEmpty(jsonSearchResult))
            {
                // Convert from JSON to a typed array
                var searchResultItems = JsonConvert.DeserializeObject<PnPTemplatesGalleryResultItem[]>(jsonSearchResult);

                result = (from r in searchResultItems
                          select new ProvisioningTemplateInformation
                          {
                              DisplayName = r.Title,
                              Description = r.Abstract,
                              TemplateFileUri = r.TemplatePnPUrl,
                              TemplateImageUrl = r.ImageUrl,
                              Scope = r.Scopes,
                              Platforms = r.Platforms,
                          }).ToArray();
            }

            return (result);
        }
    }
}
