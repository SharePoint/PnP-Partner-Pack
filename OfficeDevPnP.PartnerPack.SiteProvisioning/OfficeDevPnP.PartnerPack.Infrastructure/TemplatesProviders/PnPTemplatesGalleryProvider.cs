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

            if (configuration.Name != "gallery")
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
            HttpResponseHeaders responseHeaders; 

            // Get the template via HTTP REST
            var templateStream = HttpHelper.MakeGetRequestForStreamWithResponseHeaders(
                $"{this._templatesGalleryBaseUrl}api/DownloadTemplate/{templateUri}",
                "application/octet-stream", out responseHeaders);

            // If we have any result
            if (templateStream != null)
            {
                XMLTemplateProvider provider = new XMLOpenXMLTemplateProvider(
                    new OpenXMLConnector(templateStream));

                // Read the .PNP Open XML file name
                if (responseHeaders.Contains("Content-Disposition"))
                {
                    // Read the content disposition header value
                    var contentDispositionHeader = responseHeaders.FirstOrDefault(h => h.Key == "Content-Disposition");
                    var contentDisposition = new System.Net.Mime.ContentDisposition(contentDispositionHeader.Value.First());

                    var openXMLFileName = contentDisposition.FileName;

                    // Determine the name of the XML file inside the PNP Open XML file
                    var xmlTemplateFile = openXMLFileName.ToLower().Replace(".pnp", ".xml");

                    // Get the template
                    result = provider.GetTemplate(xmlTemplateFile);
                }
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

            String targetPlatforms = String.Empty;
            String targetScopes = String.Empty;

            // Search via HTTP REST
            var jsonSearchResult = HttpHelper.MakeGetRequestForString(
                $"{this._templatesGalleryBaseUrl}api/SearchTemplates?searchText={searchText}&platforms={targetPlatforms}&scopes={targetScopes}");

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
