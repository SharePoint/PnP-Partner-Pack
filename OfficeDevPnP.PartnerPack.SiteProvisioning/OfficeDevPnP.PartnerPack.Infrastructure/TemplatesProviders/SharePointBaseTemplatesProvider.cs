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

namespace OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders
{
    /// <summary>
    /// Base abstract class for any SharePoint-based templates provider
    /// </summary>
    public abstract class SharePointBaseTemplatesProvider : ITemplatesProvider
    {
        public String TemplatesSiteUrl { get; protected set; }

        public abstract string DisplayName { get; }

        public SharePointBaseTemplatesProvider()
        {
            // If the ParentSiteUrl is empty or NULL, then fallback to Global Tenant settings
            this.TemplatesSiteUrl = !String.IsNullOrEmpty(PnPPartnerPackSettings.ParentSiteUrl) ?
                PnPPartnerPackSettings.ParentSiteUrl :
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

                var templateRelativePath = templateUri.Substring(0, web.Url.Length + 1);
                templateRelativePath = templateUri.Substring(0, PnPPartnerPackConstants.PnPProvisioningTemplates.Length + 1);

                // Configure the SharePoint Connector
                var sharepointConnector = new SharePointConnector(context, web.Url,
                    PnPPartnerPackConstants.PnPProvisioningTemplates);

                // Configure the Open XML provider
                XMLTemplateProvider provider =
                    new XMLOpenXMLTemplateProvider(
                        new OpenXMLConnector(templateRelativePath, sharepointConnector));

                // Determine the name of the XML file inside the PNP Open XML file
                var xmlTemplateFile = templateRelativePath.ToLower().Replace(".pnp", ".xml");

                // Get the template
                ProvisioningTemplate template = provider.GetTemplate(xmlTemplateFile);

                return (template);
            }
        }

        public virtual ProvisioningTemplateInformation[] SearchProvisioningTemplates(string searchText, TargetPlatform platforms, TargetScope scope)
        {
            List<ProvisioningTemplateInformation> result =
                new List<ProvisioningTemplateInformation>();

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
                        platformsCAMLFilter = @"<Contains>
                                                    <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                    <Value Type='Text'>SharePoint Online</Value>
                                                </Contains>";
                    }
                    if ((platforms & TargetPlatform.SharePoint2016) == TargetPlatform.SharePoint2016)
                    {
                        if (!String.IsNullOrEmpty(platformsCAMLFilter))
                        {
                            platformsCAMLFilter = @"<Or>" +
                                                        platformsCAMLFilter + @"
                                                        <Contains>
                                                            <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                            <Value Type='Text'>SharePoint 2016</Value>
                                                        </Contains>
                                                    </Or>";
                        }
                        else
                        {
                            platformsCAMLFilter = @"<Contains>
                                                    <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                    <Value Type='Text'>SharePoint 2016</Value>
                                                </Contains>";
                        }
                    }
                    if ((platforms & TargetPlatform.SharePoint2013) == TargetPlatform.SharePoint2013)
                    {
                        if (!String.IsNullOrEmpty(platformsCAMLFilter))
                        {
                            platformsCAMLFilter = @"<Or>" +
                                                        platformsCAMLFilter + @"
                                                        <Contains>
                                                            <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                            <Value Type='Text'>SharePoint 2013</Value>
                                                        </Contains>
                                                    </Or>";
                        }
                        else
                        {
                            platformsCAMLFilter = @"<Contains>
                                                    <FieldRef Name='PnPProvisioningTemplatePlatform' />
                                                    <Value Type='Text'>SharePoint 2013</Value>
                                                </Contains>";
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
                                (!String.IsNullOrEmpty(platformsCAMLFilter) ? "<And>" : String.Empty) + @"
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

                            // Configure the Open XML provider
                            XMLTemplateProvider provider =
                                new XMLOpenXMLTemplateProvider(
                                    new OpenXMLConnector(item.File.Name, sharepointConnector));

                            // Determine the name of the XML file inside the PNP Open XML file
                            var xmlTemplateFile = item.File.Name.ToLower().Replace(".pnp", ".xml");

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

                            result.Add(templateInformation);
                        }
                    }
                    catch (ServerException)
                    {
                        // In case of any issue, ignore the failing templates
                    }
                }

                return (result.ToArray());
            }
        }
    }
}
