using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;
using System.Runtime.Caching;
using Newtonsoft.Json;

namespace OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders
{
    /// <summary>
    /// Base abstract class for any SharePoint-based templates provider
    /// </summary>
    public abstract class SharePointBaseTemplatesProvider : ITemplatesProvider
    {
        // Lazy static private property for managing the in memory cache
        private static Lazy<MemoryCache> cacheValue = new Lazy<MemoryCache>(() =>
        {
            MemoryCache result = new MemoryCache("SharePointTemplatesProvider");
            return (result);
        }, true);

        protected static MemoryCache Cache
        {
            get
            {
                return (cacheValue.Value);
            }
        }

        public String TemplatesSiteUrl { get; protected set; }

        public abstract string DisplayName { get; }

        public SharePointBaseTemplatesProvider()
        {
        }

        public SharePointBaseTemplatesProvider(String templateSiteUrl)
        {
            // If the ParentSiteUrl is empty or NULL, then fallback to Global Tenant settings
            this.TemplatesSiteUrl = !String.IsNullOrEmpty(templateSiteUrl) ?
                templateSiteUrl :
                PnPPartnerPackSettings.InfrastructureSiteUrl;
        }

        public void Init(XElement configuration)
        {
        }

        public virtual ProvisioningTemplate GetProvisioningTemplate(string templateUri)
        {
            // Connect to the target Templates Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(TemplatesSiteUrl))
            {
                // Get a reference to the target library
                Web web = context.Web;

                web.EnsureProperty(w => w.Url);

                var templateRelativePath = templateUri.Substring(templateUri.LastIndexOf("/") + 1);

                // Configure the SharePoint Connector
                var sharepointConnector = new SharePointConnector(context, web.Url,
                    PnPPartnerPackConstants.PnPProvisioningTemplates);

                TemplateProviderBase provider = null;
                // If the target is a .PNP Open XML template
                if (templateRelativePath.ToLower().EndsWith(".pnp"))
                {
                    // Configure the Open XML provider for SharePoint
                    provider =
                        new XMLOpenXMLTemplateProvider(
                            new OpenXMLConnector(templateRelativePath, sharepointConnector));
                }
                else
                {
                    // Otherwise use the .XML template provider for SharePoint
                    provider =
                        new XMLSharePointTemplateProvider(context, web.Url,
                            PnPPartnerPackConstants.PnPProvisioningTemplates);
                }

                // Determine the name of the XML file inside the PNP Open XML file, if any
                var xmlTemplateFile = templateRelativePath.ToLower().Replace(".pnp", ".xml");

                // Get the template
                ProvisioningTemplate template = provider.GetTemplate(xmlTemplateFile);
                template.Connector = provider.Connector;

                return (template);
            }
        }

        public virtual ProvisioningTemplateInformation[] SearchProvisioningTemplates(string searchText, TargetPlatform platforms, TargetScope scope)
        {
            String cacheKey = JsonConvert.SerializeObject(new SharePointSearchCacheKey
            {
                TemplatesProviderTypeName = this.GetType().Name,
                SearchText = searchText,
                Platforms = platforms,
                Scope = scope,
            });

            List<ProvisioningTemplateInformation> result = Cache[cacheKey] as List<ProvisioningTemplateInformation>;

            if (result == null)
            {
                result = SearchProvisioningTemplatesInternal(searchText, platforms, scope, cacheKey);
            }

            return (result.ToArray());
        }

        private List<ProvisioningTemplateInformation> SearchProvisioningTemplatesInternal(string searchText, TargetPlatform platforms, TargetScope scope, String cacheKey)
        {
            List<ProvisioningTemplateInformation> result = new List<ProvisioningTemplateInformation>();

            // Connect to the target Templates Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(TemplatesSiteUrl))
            {
                // Get a reference to the target library
                Web web = context.Web;

                String platformsCAMLFilter = null;

                // Build the target Platforms filter
                if (platforms != TargetPlatform.None && platforms != TargetPlatform.All)
                {
                    if ((platforms & TargetPlatform.SharePointOnline) == TargetPlatform.SharePointOnline)
                    {
                        platformsCAMLFilter = @"<Eq>
                                                    <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                    <Value Type='MultiChoice'>SharePoint Online</Value>
                                                </Eq>";
                    }
                    if ((platforms & TargetPlatform.SharePoint2016) == TargetPlatform.SharePoint2016)
                    {
                        if (!String.IsNullOrEmpty(platformsCAMLFilter))
                        {
                            platformsCAMLFilter = @"<Or>" +
                                                        platformsCAMLFilter + @"
                                                        <Eq>
                                                            <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                            <Value Type='MultiChoice'>SharePoint 2016</Value>
                                                        </Eq>
                                                    </Or>";
                        }
                        else
                        {
                            platformsCAMLFilter = @"<Eq>
                                                    <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                    <Value Type='MultiChoice'>SharePoint 2016</Value>
                                                </Eq>";
                        }
                    }
                    if ((platforms & TargetPlatform.SharePoint2013) == TargetPlatform.SharePoint2013)
                    {
                        if (!String.IsNullOrEmpty(platformsCAMLFilter))
                        {
                            platformsCAMLFilter = @"<Or>" +
                                                        platformsCAMLFilter + @"
                                                        <Eq>
                                                            <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                            <Value Type='MultiChoice'>SharePoint 2013</Value>
                                                        </Eq>
                                                    </Or>";
                        }
                        else
                        {
                            platformsCAMLFilter = @"<Eq>
                                                    <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                    <Value Type='MultiChoice'>SharePoint 2013</Value>
                                                </Eq>";
                        }
                    }

                    try
                    {
                        List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningTemplates);

                        // Get only Provisioning Templates documents with the specified Scope
                        CamlQuery query = new CamlQuery();
                        query.ViewXml =
                            @"<View>
                        <Query>
                            <Where>" +
                                (!String.IsNullOrEmpty(platformsCAMLFilter) ? " < And>" : String.Empty) + @"
                                    <And>
                                        <Eq>
                                            <FieldRef Name='PnPProvisioningTemplateScope' />
                                            <Value Type='Choice'>" + scope.ToString() + @"</Value>
                                        </Eq>
                                        <Eq>
                                            <FieldRef Name='ContentType' />
                                            <Value Type=''Computed''>PnPProvisioningTemplate</Value>
                                        </Eq>
                                    </And>" + platformsCAMLFilter +
                                (!String.IsNullOrEmpty(platformsCAMLFilter) ? "</And>" : String.Empty) + @"                                
                            </Where>
                        </Query>
                        <ViewFields>
                            <FieldRef Name='Title' />
                            <FieldRef Name='PnPProvisioningTemplateScope' />
                            <FieldRef Name='PnPProvisioningTemplatePlatform' />
                            <FieldRef Name='PnPProvisioningTemplateSourceUrl' />
                        </ViewFields>
                    </View>";

                        ListItemCollection items = list.GetItems(query);
                        context.Load(items,
                            includes => includes.Include(i => i.File,
                            i => i[PnPPartnerPackConstants.PnPProvisioningTemplateScope],
                            i => i[PnPPartnerPackConstants.PnPProvisioningTemplatePlatform],
                            i => i[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl]));
                        context.ExecuteQueryRetry();

                        web.EnsureProperty(w => w.Url);

                        // Configure the SharePoint Connector
                        var sharepointConnector = new SharePointConnector(context, web.Url,
                                PnPPartnerPackConstants.PnPProvisioningTemplates);

                        foreach (ListItem item in items)
                        {
                            // Get the template file name and server relative URL
                            item.File.EnsureProperties(f => f.Name, f => f.ServerRelativeUrl);

                            TemplateProviderBase provider = null;

                            // If the target is a .PNP Open XML template
                            if (item.File.Name.ToLower().EndsWith(".pnp"))
                            {
                                // Configure the Open XML provider for SharePoint
                                provider =
                                    new XMLOpenXMLTemplateProvider(
                                        new OpenXMLConnector(item.File.Name, sharepointConnector));
                            }
                            else
                            {
                                // Otherwise use the .XML template provider for SharePoint
                                provider =
                                    new XMLSharePointTemplateProvider(context, web.Url,
                                        PnPPartnerPackConstants.PnPProvisioningTemplates);
                            }

                            // Determine the name of the XML file inside the PNP Open XML file, if any
                            var xmlTemplateFile = item.File.Name.ToLower().Replace(".pnp", ".xml");

                            try
                            {
                                // Get the template
                                ProvisioningTemplate template = provider.GetTemplate(xmlTemplateFile);

                                // Prepare the resulting item
                                var templateInformation = new ProvisioningTemplateInformation
                                {
                                    // Scope = (TargetScope)Enum.Parse(typeof(TargetScope), (String)item[PnPPartnerPackConstants.PnPProvisioningTemplateScope], true),
                                    TemplateSourceUrl = item[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl] != null ? ((FieldUrlValue)item[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl]).Url : null,
                                    TemplateFileUri = String.Format("{0}/{1}/{2}", web.Url, PnPPartnerPackConstants.PnPProvisioningTemplates, item.File.Name),
                                    TemplateImageUrl = template.ImagePreviewUrl,
                                    DisplayName = template.DisplayName,
                                    Description = template.Description,
                                };

                                #region Determine Scope

                                String targetScope;
                                if (template.Properties.TryGetValue(PnPPartnerPackConstants.TEMPLATE_SCOPE, out targetScope))
                                {
                                    if (String.Equals(targetScope, PnPPartnerPackConstants.TEMPLATE_SCOPE_PARTIAL, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        templateInformation.Scope = TargetScope.Partial;
                                    }
                                    else if (String.Equals(targetScope, PnPPartnerPackConstants.TEMPLATE_SCOPE_WEB, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        templateInformation.Scope = TargetScope.Web;
                                    }
                                    else if (String.Equals(targetScope, PnPPartnerPackConstants.TEMPLATE_SCOPE_SITE, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        templateInformation.Scope = TargetScope.Site;
                                    }
                                }

                                #endregion

                                #region Determine target Platforms

                                String spoPlatform, sp2016Platform, sp2013Platform;
                                if (template.Properties.TryGetValue(PnPPartnerPackConstants.PLATFORM_SPO, out spoPlatform))
                                {
                                    if (spoPlatform.Equals(PnPPartnerPackConstants.TRUE_VALUE, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        templateInformation.Platforms |= TargetPlatform.SharePointOnline;
                                    }
                                }
                                if (template.Properties.TryGetValue(PnPPartnerPackConstants.PLATFORM_SP2016, out sp2016Platform))
                                {
                                    if (sp2016Platform.Equals(PnPPartnerPackConstants.TRUE_VALUE, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        templateInformation.Platforms |= TargetPlatform.SharePoint2016;
                                    }
                                }
                                if (template.Properties.TryGetValue(PnPPartnerPackConstants.PLATFORM_SP2013, out sp2013Platform))
                                {
                                    if (sp2013Platform.Equals(PnPPartnerPackConstants.TRUE_VALUE, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        templateInformation.Platforms |= TargetPlatform.SharePoint2013;
                                    }
                                }

                                #endregion

                                // If we don't have a search text 
                                // or we have a search text and it is contained either 
                                // in the DisplayName or in the Description of the template
                                if ((!String.IsNullOrEmpty(searchText) &&
                                    ((!String.IsNullOrEmpty(template.DisplayName) && template.DisplayName.ToLower().Contains(searchText.ToLower())) ||
                                    (!String.IsNullOrEmpty(template.Description) && template.Description.ToLower().Contains(searchText.ToLower())))) ||
                                    String.IsNullOrEmpty(searchText))
                                {
                                    // Add the template to the result
                                    result.Add(templateInformation);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Ignore any exception related to the current template
                                // and move to the next template
                            }
                        }
                    }
                    catch (ServerException)
                    {
                        // In case of any issue, ignore the failing templates
                    }
                }
            }

            CacheItemPolicy policy = new CacheItemPolicy
            {
                Priority = CacheItemPriority.Default,
                AbsoluteExpiration = DateTimeOffset.Now.AddMinutes(30), // Cache results for 30 minutes
                RemovedCallback = (args) =>
                {
                    if (args.RemovedReason == CacheEntryRemovedReason.Expired)
                    {
                        var removedKey = args.CacheItem.Key;
                        var searchInputs = JsonConvert.DeserializeObject<SharePointSearchCacheKey>(removedKey);

                        var newItem = SearchProvisioningTemplatesInternal(
                            searchInputs.SearchText,
                            searchInputs.Platforms,
                            searchInputs.Scope,
                            removedKey);
                    }
                },
            };

            Cache.Set(cacheKey, result, policy);

            return result;
        }
    }
}
