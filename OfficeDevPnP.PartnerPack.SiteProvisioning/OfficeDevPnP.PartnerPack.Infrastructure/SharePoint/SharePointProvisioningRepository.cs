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

                foreach (ListItem item in items)
                {
                    result.Add(new ProvisioningTemplateInformation
                    {
                        Scope = (TemplateScope)Enum.Parse(typeof(TemplateScope), (String)item[PnPPartnerPackConstants.PnPProvisioningTemplateScope], true),
                        TemplateSourceUrl = (String)item[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl],
                        TemplateFileUri = item.File.ServerRelativeUrl,
                        TemplateImageUrl = GetImageUrlFromTemplate(context, web, item.File.ServerRelativeUrl)
                    });
                }
            }

            return (result.ToArray());
        }

        private String GetImageUrlFromTemplate(ClientContext context, Web web, String fileServerRelativeUrl)
        {
            // Configure the XML file system provider
            XMLTemplateProvider provider =
                new XMLSharePointTemplateProvider(context, web.Url,
                    PnPPartnerPackConstants.PnPProvisioningTemplates);

            ProvisioningTemplate template = provider.GetTemplate(fileServerRelativeUrl);

            return (template.ImagePreviewUrl);
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

            // TODO: Implement this one
            templateToSave.ImagePreviewUrl = "fake.png";

            // And save it on the file system
            provider.SaveAs(templateToSave, job.FileName);

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
                if (items.Count > 0)
                {
                    ListItem jobItem = items[0];
                    return (PrepareJobInformationFromSharePoint(jobItem));
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
                foreach (var jobItem in items)
                {
                    result.Add(PrepareJobInformationFromSharePoint(jobItem));
                }
            }
            return (result.ToArray());
        }

        private static ProvisioningJobInformation PrepareJobInformationFromSharePoint(ListItem jobItem)
        {
            ProvisioningJobInformation resultItem = new ProvisioningJobInformation();
            resultItem.JobId = Guid.Parse(((String)jobItem["LinkFilename"]).Substring(0, ((String)jobItem["LinkFilename"]).Length - 4));
            resultItem.Title = (String)jobItem[PnPPartnerPackConstants.TitleField];
            resultItem.Status = (ProvisioningJobStatus)Enum.Parse(typeof(ProvisioningJobStatus), (String)jobItem[PnPPartnerPackConstants.PnPProvisioningJobStatus]);
            resultItem.ErrorMessage = (String)jobItem[PnPPartnerPackConstants.PnPProvisioningJobError];
            resultItem.Type = (String)jobItem[PnPPartnerPackConstants.PnPProvisioningJobType];
            resultItem.Owner = ((FieldUserValue)jobItem[PnPPartnerPackConstants.PnPProvisioningJobOwner]).LookupValue;
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
