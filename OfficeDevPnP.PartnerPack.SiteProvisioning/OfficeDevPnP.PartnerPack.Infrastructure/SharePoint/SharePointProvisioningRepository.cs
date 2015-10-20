using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.SharePoint
{
    /// <summary>
    /// Provides the implementation of a Provisioning Repository that targets SharePoint
    /// </summary>
    public class SharePointProvisioningRepository : IProvisioningRepository
    {
        public void Init()
        {
            // NOOP
            return;
        }

        public ProvisioningTemplateInformation[] GetGlobalProvisioningTemplates(TemplateScope scope)
        {
            return (GetLocalProvisioningTemplates(PnPPartnerPackSettings.InfrastructureSiteUrl, scope));
        }

        public ProvisioningTemplateInformation[] GetLocalProvisioningTemplates(string siteUrl, TemplateScope scope)
        {
            List<ProvisioningTemplateInformation> result =
                new List<ProvisioningTemplateInformation>();

            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                // Get a reference to the target library
                Web web = context.Web;

                try
                {
                    List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningTemplates);

                    // Get only Provisioning Templates documents with the specified Scope
                    CamlQuery query = new CamlQuery();
                    query.ViewXml =
                        @"<View>
                        <Query>
                            <Where>
                                <And>
                                    <Eq>
                                        <FieldRef Name='PnPProvisioningTemplateScope' />
                                        <Value Type='Choice'>" + scope.ToString() + @"</Value>
                                    </Eq>
                                    <Eq>
                                        <FieldRef Name='ContentType' />
                                        <Value Type=''Computed''>PnPProvisioningTemplate</Value>
                                    </Eq>
                                </And>
                            </Where>
                        </Query>
                        <ViewFields>
                            <FieldRef Name='Title' />
                            <FieldRef Name='PnPProvisioningTemplateScope' />
                            <FieldRef Name='PnPProvisioningTemplateSourceUrl' />
                        </ViewFields>
                    </View>";

                    ListItemCollection items = list.GetItems(query);
                    context.Load(items,
                        includes => includes.Include(i => i.File,
                        i => i[PnPPartnerPackConstants.PnPProvisioningTemplateScope],
                        i => i[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl]));
                    context.ExecuteQueryRetry();

                    web.EnsureProperty(w => w.Url);

                    foreach (ListItem item in items)
                    {
                        // Configure the XML file system provider
                        XMLTemplateProvider provider =
                            new XMLSharePointTemplateProvider(context, web.Url,
                                PnPPartnerPackConstants.PnPProvisioningTemplates);

                        item.File.EnsureProperties(f => f.Name, f => f.ServerRelativeUrl);

                        ProvisioningTemplate template = provider.GetTemplate(item.File.Name);

                        result.Add(new ProvisioningTemplateInformation
                        {
                            Scope = (TemplateScope)Enum.Parse(typeof(TemplateScope), (String)item[PnPPartnerPackConstants.PnPProvisioningTemplateScope], true),
                            TemplateSourceUrl = ((FieldUrlValue)item[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl]).Url,
                            TemplateFileUri = String.Format("{0}/{1}/{2}", web.Url, PnPPartnerPackConstants.PnPProvisioningTemplates, item.File.Name),
                            TemplateImageUrl = template.ImagePreviewUrl,
                            DisplayName = template.DisplayName,
                            Description = template.Description,
                        });
                    }
                }
                catch (ServerException ex)
                {
                    // In case of any issue, ignore the local templates
                }
            }

            return (result.ToArray());
        }

        /// <summary>
        /// Saves a Provisioning Template into the target Global repository
        /// </summary>
        /// <param name="job">The Provisioning Template to save</param>
        /// <returns>The ID  of the saved Provisioning Template</returns>
        public Guid SaveGlobalProvisioningTemplate(GetProvisioningTemplateJob job)
        {
            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                return (SaveProvisioningTemplateInternal(context, job));
            }
        }

        /// <summary>
        /// Saves a Provisioning Template into the target Local repository
        /// </summary>
        /// <param name="siteUrl">The local Site Collection to save to</param>
        /// <param name="template">The Provisioning Template to save</param>
        /// <returns>The ID  of the saved Provisioning Template</returns>
        public Guid SaveLocalProvisioningTemplate(string siteUrl, GetProvisioningTemplateJob job)
        {
            // Connect to the Local Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                PnPPartnerPackUtilities.EnablePartnerPackInfrastructureOnSite(siteUrl);
                return (SaveProvisioningTemplateInternal(context, job));
            }
        }

        private Guid SaveProvisioningTemplateInternal(ClientContext context, GetProvisioningTemplateJob job)
        {
            // Get a reference to the target web site
            Web web = context.Web;
            context.Load(web, w => w.Url);
            context.ExecuteQueryRetry();

            // Configure the XML file system provider
            XMLTemplateProvider provider =
                new XMLSharePointTemplateProvider(context, web.Url,
                    PnPPartnerPackConstants.PnPProvisioningTemplates);

            ProvisioningTemplateCreationInformation ptci =
                new ProvisioningTemplateCreationInformation(web);
            ptci.FileConnector = provider.Connector;
            ptci.IncludeAllTermGroups = job.IncludeAllTermGroups;
            ptci.IncludeSearchConfiguration = job.IncludeSearchConfiguration;
            ptci.IncludeSiteCollectionTermGroup = job.IncludeSiteCollectionTermGroup;
            ptci.IncludeSiteGroups = job.IncludeSiteGroups;
            ptci.PersistComposedLookFiles = job.PersistComposedLookFiles;

            // Extract the current template
            ProvisioningTemplate templateToSave = web.GetProvisioningTemplate(ptci);

            templateToSave.Description = job.Description;
            templateToSave.DisplayName = job.Title;

            // Save template image preview in folder
            Folder templatesFolder = web.GetFolderByServerRelativeUrl(PnPPartnerPackConstants.PnPProvisioningTemplates);
            context.Load(templatesFolder, f => f.ServerRelativeUrl, f => f.Name);
            context.ExecuteQueryRetry();

            String previewImageFileName = job.FileName.Replace(".xml", "_preview.png");
            templatesFolder.UploadFile(previewImageFileName,
                job.TemplateImageFile.ToStream(), true);

            // And store URL in the XML file
            templateToSave.ImagePreviewUrl = String.Format("{0}/{1}/{2}",
                web.Url, templatesFolder.Name, previewImageFileName);

            // And save it on the file system
            provider.SaveAs(templateToSave, job.FileName);

            Microsoft.SharePoint.Client.File templateFile = templatesFolder.GetFile(job.FileName);
            ListItem item = templateFile.ListItemAllFields;

            item[PnPPartnerPackConstants.ContentTypeIdField] = PnPPartnerPackConstants.PnPProvisioningTemplateContentTypeId;
            item[PnPPartnerPackConstants.TitleField] = job.Title;
            item[PnPPartnerPackConstants.PnPProvisioningTemplateScope] = job.Scope.ToString();
            item[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl] = job.SourceSiteUrl;

            item.Update();

            context.ExecuteQueryRetry();

            // TODO: Replace with real ID of the file
            return (Guid.NewGuid());
        }

        public Guid EnqueueProvisioningJob(ProvisioningJob job)
        {
            // Prepare the Job ID
            Guid jobId = Guid.NewGuid();

            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Convert the current Provisioning Job into a Stream
                Stream stream = job.ToJsonStream();

                // Get a reference to the target library
                Web web = context.Web;
                List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningJobs);
                Microsoft.SharePoint.Client.File file = list.RootFolder.UploadFile(String.Format("{0}.job", jobId), stream, false);

                Microsoft.SharePoint.Client.User ownerUser = web.EnsureUser(job.Owner);
                context.Load(ownerUser);
                context.ExecuteQueryRetry();

                ListItem item = file.ListItemAllFields;

                item[PnPPartnerPackConstants.ContentTypeIdField] = PnPPartnerPackConstants.PnPProvisioningJobContentTypeId;
                item[PnPPartnerPackConstants.TitleField] = job.Title;
                item[PnPPartnerPackConstants.PnPProvisioningJobStatus] = ProvisioningJobStatus.Pending.ToString();
                item[PnPPartnerPackConstants.PnPProvisioningJobError] = String.Empty;
                item[PnPPartnerPackConstants.PnPProvisioningJobType] = job.GetType().FullName;

                FieldUserValue ownerUserValue = new FieldUserValue();
                ownerUserValue.LookupId = ownerUser.Id;
                item[PnPPartnerPackConstants.PnPProvisioningJobOwner] = ownerUserValue;

                item.Update();

                context.ExecuteQueryRetry();
            }

            return (jobId);
        }

        public ProvisioningJobInformation GetProvisioningJob(Guid jobId)
        {
            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Get a reference to the target library
                Web web = context.Web;
                List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningJobs);

                CamlQuery query = new CamlQuery();
                query.ViewXml =
                    @"<View>
                        <Query>
                            <Where>
                                <Eq>
                                    <FieldRef Name='Name' />
                                    <Value Type='Text'>" + jobId + @".job</Value>
                                </Eq>
                            </Where>
                        </Query>
                    </View>";

                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQueryRetry();

                if (items.Count > 0)
                {
                    ListItem jobItem = items[0];
                    return (PrepareJobInformationFromSharePoint(context, jobItem, true));
                }
                else
                {
                    return (null);
                }
            }
        }

        public ProvisioningJobInformation[] GetProvisioningJobs(ProvisioningJobStatus status, string owner = null)
        {
            List<ProvisioningJobInformation> result = new List<ProvisioningJobInformation>();

            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Get a reference to the target library
                Web web = context.Web;
                List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningJobs);

                CamlQuery query = new CamlQuery();
                query.ViewXml =
                    @"<View>
                        <Query>
                            <Where>
                                <Eq>
                                    <FieldRef Name='PnPProvisioningJobStatus' />
                                    <Value Type='Text'>" + status + @"</Value>
                                </Eq>
                            </Where>
                        </Query>
                    </View>";

                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQueryRetry();

                foreach (var jobItem in items)
                {
                    result.Add(PrepareJobInformationFromSharePoint(context, jobItem, true));
                }
            }
            return (result.ToArray());
        }

        private static ProvisioningJobInformation PrepareJobInformationFromSharePoint(ClientContext context, ListItem jobItem, Boolean includeFileStream = false)
        {
            ProvisioningJobInformation resultItem = new ProvisioningJobInformation();
            resultItem.JobId = Guid.Parse(((String)jobItem["FileLeafRef"]).Substring(0, ((String)jobItem["FileLeafRef"]).Length - 4));
            resultItem.Title = (String)jobItem[PnPPartnerPackConstants.TitleField];
            resultItem.Status = (ProvisioningJobStatus)Enum.Parse(typeof(ProvisioningJobStatus), (String)jobItem[PnPPartnerPackConstants.PnPProvisioningJobStatus]);
            resultItem.ErrorMessage = (String)jobItem[PnPPartnerPackConstants.PnPProvisioningJobError];
            resultItem.Type = (String)jobItem[PnPPartnerPackConstants.PnPProvisioningJobType];
            resultItem.Owner = ((FieldUserValue)jobItem[PnPPartnerPackConstants.PnPProvisioningJobOwner]).LookupValue;

            if (includeFileStream)
            {
                jobItem.ParentList.RootFolder.EnsureProperty(f => f.ServerRelativeUrl);

                Microsoft.SharePoint.Client.File jobFile = jobItem.ParentList.ParentWeb.GetFileByServerRelativeUrl(
                    String.Format("{0}/{1}", jobItem.ParentList.RootFolder.ServerRelativeUrl, (String)jobItem["FileLeafRef"]));
                context.Load(jobFile, jf => jf.ServerRelativeUrl);

                var jobFileStream = jobFile.OpenBinaryStream();
                context.ExecuteQueryRetry();

                resultItem.JobServerRelativeUrl = jobFile.ServerRelativeUrl;

                MemoryStream mem = new MemoryStream();
                jobFileStream.Value.CopyTo(mem);
                mem.Position = 0;
                resultItem.JobFile = mem;
            }

            return resultItem;
        }

        public void UpdateProvisioningJob(Guid jobId, ProvisioningJobStatus status, String errorMessage = null)
        {
            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Get a reference to the target library
                Web web = context.Web;
                List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningJobs);

                CamlQuery query = new CamlQuery();
                query.ViewXml =
                    @"<View>
                        <Query>
                            <Where>
                                <Eq>
                                    <FieldRef Name='Name' />
                                    <Value Type='Text'>" + jobId + @".job</Value>
                                </Eq>
                            </Where>
                        </Query>
                    </View>";

                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQueryRetry();

                if (items.Count > 0)
                {
                    ListItem jobItem = items[0];

                    jobItem[PnPPartnerPackConstants.PnPProvisioningJobStatus] = status;
                    jobItem[PnPPartnerPackConstants.PnPProvisioningJobError] = errorMessage;

                    jobItem.Update();
                    context.ExecuteQueryRetry();
                }
            }
        }
    }
}
