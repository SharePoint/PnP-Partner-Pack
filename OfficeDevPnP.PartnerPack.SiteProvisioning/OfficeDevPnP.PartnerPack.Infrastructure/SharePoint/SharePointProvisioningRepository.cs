using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace OfficeDevPnP.PartnerPack.Infrastructure.SharePoint
{
    /// <summary>
    /// Provides the implementation of a Provisioning Repository that targets SharePoint
    /// </summary>
    public class SharePointProvisioningRepository : IProvisioningRepository
    {
        public void Init(XElement configuration)
        {
            // NOOP
            return;
        }

        public ProvisioningTemplateInformation[] GetGlobalProvisioningTemplates(TargetScope scope)
        {
            return (GetLocalProvisioningTemplates(PnPPartnerPackSettings.InfrastructureSiteUrl, scope));
        }

        public ProvisioningTemplateInformation[] GetLocalProvisioningTemplates(string siteUrl, TargetScope scope)
        {
            List<ProvisioningTemplateInformation> result =
                new List<ProvisioningTemplateInformation>();

            // Retrieve the Root Site Collection URL
            siteUrl = PnPPartnerPackUtilities.GetSiteCollectionRootUrl(siteUrl);

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
                            Scope = (TargetScope)Enum.Parse(typeof(TargetScope), (String)item[PnPPartnerPackConstants.PnPProvisioningTemplateScope], true),
                            TemplateSourceUrl = item[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl] != null ? ((FieldUrlValue)item[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl]).Url : null,
                            TemplateFileUri = String.Format("{0}/{1}/{2}", web.Url, PnPPartnerPackConstants.PnPProvisioningTemplates, item.File.Name),
                            TemplateImageUrl = template.ImagePreviewUrl,
                            DisplayName = template.DisplayName,
                            Description = template.Description,
                        });
                    }
                }
                catch (ServerException)
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
        public void SaveGlobalProvisioningTemplate(GetProvisioningTemplateJob job)
        {
            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(job.SourceSiteUrl))
            {
                SaveProvisioningTemplateInternal(context, job, true);
            }
        }

        /// <summary>
        /// Saves a Provisioning Template into the target Local repository
        /// </summary>
        /// <param name="siteUrl">The local Site Collection to save to</param>
        /// <param name="template">The Provisioning Template to save</param>
        public void SaveLocalProvisioningTemplate(string siteUrl, GetProvisioningTemplateJob job)
        {
            // Connect to the Local Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(siteUrl))
            {
                PnPPartnerPackUtilities.EnablePartnerPackInfrastructureOnSite(siteUrl);
                SaveProvisioningTemplateInternal(context, job, false);
            }
        }

        private void SaveProvisioningTemplateInternal(ClientContext context, GetProvisioningTemplateJob job, Boolean globalRepository = true)
        {
            // Fix the job filename if it is missing the .pnp extension
            if (!job.FileName.ToLower().EndsWith(".pnp"))
            {
                job.FileName += ".pnp";
            }

            // Get the Access Token from the current context
            var accessToken = context.GetAccessToken();

            // Get a reference to the target web site
            Web web = context.Web;
            context.Load(web, w => w.Url, w => w.ServerRelativeUrl);
            context.ExecuteQueryRetry();

            // Prepare the support variables
            ClientContext repositoryContext = null;
            Web repositoryWeb = null;
            
            // Define whether we need to use the global infrastructural repository or the local one
            if (globalRepository)
            {
                // Get a reference to the global repository web site and context
                repositoryContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(
                    PnPPartnerPackSettings.InfrastructureSiteUrl);
            }
            else
            {
                // Get a reference to the local repository web site and context
                repositoryContext = web.Context.GetSiteCollectionContext();
            }

            using (repositoryContext)
            {
                repositoryWeb = repositoryContext.Site.RootWeb;
                repositoryContext.Load(repositoryWeb, w => w.Url);
                repositoryContext.ExecuteQueryRetry();

                // Configure the XML SharePoint provider for the Infrastructural Site Collection
                XMLTemplateProvider provider = new XMLOpenXMLTemplateProvider(job.FileName,
                    new SharePointConnector(repositoryContext, repositoryWeb.Url,
                            PnPPartnerPackConstants.PnPProvisioningTemplates));

                ProvisioningTemplateCreationInformation ptci =
                    new ProvisioningTemplateCreationInformation(web);
                ptci.FileConnector = provider.Connector;
                ptci.IncludeAllTermGroups = job.IncludeAllTermGroups;
                ptci.IncludeSearchConfiguration = job.IncludeSearchConfiguration;
                ptci.IncludeSiteCollectionTermGroup = job.IncludeSiteCollectionTermGroup;
                ptci.IncludeSiteGroups = job.IncludeSiteGroups;
                ptci.PersistBrandingFiles = job.PersistComposedLookFiles;

                // We do intentionally remove taxonomies and search, which are not supported 
                // in the AppOnly Authorization model
                // For further details, see the PnP Partner Pack documentation 
                ptci.HandlersToProcess ^= Handlers.TermGroups;
                ptci.HandlersToProcess ^= Handlers.SearchSettings;

                // Extract the current template
                ProvisioningTemplate templateToSave = web.GetProvisioningTemplate(ptci);

                templateToSave.Description = job.Description;
                templateToSave.DisplayName = job.Title;

                if (job.PersistComposedLookFiles)
                {
                    templateToSave.ComposedLook.Name = "SharePointBranding";

                    var refererUri = new Uri(web.Url);
                    var refererValue = $"{refererUri.Scheme}://{refererUri.Host}/";

                    // Create Theme Entity object
                    ThemeEntity cl = web.GetCurrentComposedLook();
                
                    if (!String.IsNullOrEmpty(cl.Font))
                    {
                        // Construct files endpoint for spfont
                        string strUrl = String.Format("{0}/_api/web/getfilebyserverrelativeurl('{1}{2}')/$value", web.Url, web.ServerRelativeUrl,cl.Font);

                        var fontFileName = cl.Font.Substring(cl.Font.LastIndexOf("/") + 1);
                        var fontFileStream = HttpHelper.MakeGetRequestForStream(strUrl, "application/octet-stream", accessToken, referer: refererValue);

                        templateToSave.ComposedLook.FontFile = String.Format("{{themecatalog}}/15/{0}", fontFileName);
                        provider.Connector.SaveFileStream(fontFileName, fontFileStream);

                        templateToSave.Files.Add(new Core.Framework.Provisioning.Model.File
                        {
                            Src = fontFileName,
                            Folder = "{themecatalog}/15",
                            Overwrite = true,
                        });
                    }
                    else
                    {
                        templateToSave.ComposedLook.FontFile = String.Empty;
                    }

                    if (!String.IsNullOrEmpty(cl.Theme))
                    {
                        // Construct files endpoint for spcolor
                        string strUrl = String.Format("{0}/_api/web/getfilebyserverrelativeurl('{1}{2}')/$value", web.Url, web.ServerRelativeUrl, cl.Theme);

                        var colorFileName = cl.Theme.Substring(cl.Theme.LastIndexOf("/") + 1);
                        var colorFileStream = HttpHelper.MakeGetRequestForStream(strUrl, "application/xhtml+xml", accessToken, referer: refererValue);

                        templateToSave.ComposedLook.ColorFile = String.Format("{{themecatalog}}/15/{0}", colorFileName);
                        provider.Connector.SaveFileStream(colorFileName, colorFileStream);

                        templateToSave.Files.Add(new Core.Framework.Provisioning.Model.File
                        {
                            Src = colorFileName,
                            Folder = "{themecatalog}/15",
                            Overwrite = true,
                        });
                    }
                    else
                    {
                        templateToSave.ComposedLook.ColorFile = String.Empty;
                    }
                }

                // Save template image preview in folder
                Microsoft.SharePoint.Client.Folder templatesFolder = repositoryWeb.GetFolderByServerRelativeUrl(PnPPartnerPackConstants.PnPProvisioningTemplates);
                repositoryContext.Load(templatesFolder, f => f.ServerRelativeUrl, f => f.Name);
                repositoryContext.ExecuteQueryRetry();

                // If there is a preview image
                if (job.TemplateImageFile != null)
                {
                    // Determine the preview image file name
                    String previewImageFileName = job.FileName.ToLower().Replace(".pnp", "_preview.png");

                    // Save the preview image inside the Open XML package
                    provider.Connector.SaveFileStream(previewImageFileName, job.TemplateImageFile.ToStream());

                    // And store URL in the XML file
                    templateToSave.ImagePreviewUrl = String.Format("{0}{1}/{2}/{3}/{4}",
                        repositoryWeb.Url.ToLower().StartsWith("https") ? "pnps" : "pnp",
                        repositoryWeb.Url.Substring(repositoryWeb.Url.IndexOf("://")), 
                        templatesFolder.Name, job.FileName, previewImageFileName);

                }

                // And save it on the file system
                provider.SaveAs(templateToSave, job.FileName.ToLower().Replace(".pnp", ".xml"));

                Microsoft.SharePoint.Client.File templateFile = templatesFolder.GetFile(job.FileName);
                ListItem item = templateFile.ListItemAllFields;

                item[PnPPartnerPackConstants.ContentTypeIdField] = PnPPartnerPackConstants.PnPProvisioningTemplateContentTypeId;
                item[PnPPartnerPackConstants.TitleField] = job.Title;
                item[PnPPartnerPackConstants.PnPProvisioningTemplateScope] = job.Scope.ToString();
                item[PnPPartnerPackConstants.PnPProvisioningTemplateSourceUrl] = job.SourceSiteUrl;

                item.Update();

                repositoryContext.ExecuteQueryRetry();
            }
        }

        public Guid EnqueueProvisioningJob(ProvisioningJob job)
        {
            // Prepare the Job ID
            Guid jobId = Guid.NewGuid();

            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Set the initial status of the Job
                job.JobId = jobId;
                job.Status = ProvisioningJobStatus.Pending;

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

                // Check if we need to enqueue a message in the Azure Storage Queue, as well
                // This happens when the Provisioning Job has to be executed in Continous mode
                if (PnPPartnerPackSettings.ContinousJobHandlers.ContainsKey(job.GetType()))
                {
                    // Get the storage account for Azure Storage Queue
                    CloudStorageAccount storageAccount =
                        CloudStorageAccount.Parse(PnPPartnerPackSettings.StorageQueueConnectionString);

                    // Get queue ... and create if it does not exist
                    CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                    CloudQueue queue = queueClient.GetQueueReference(PnPPartnerPackSettings.StorageQueueName);
                    queue.CreateIfNotExists();

                    // Add entry to queue
                    ContinousJobItem content = new ContinousJobItem { JobId = job.JobId };
                    queue.AddMessage(new CloudQueueMessage(JsonConvert.SerializeObject(content)));
                }
            }

            return (jobId);
        }

        public ProvisioningJobInformation GetProvisioningJob(Guid jobId, Boolean includeStream = false)
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
                                    <FieldRef Name='FileLeafRef' />
                                    <Value Type='Text'>" + jobId.ToString("D") + @".job</Value>
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
                    return (PrepareJobInformationFromSharePoint(context, jobItem, includeStream));
                }
                else
                {
                    return (null);
                }
            }
        }

        public ProvisioningJobInformation[] GetProvisioningJobs(ProvisioningJobStatus status, String jobType = null, Boolean includeStream = false, string owner = null)
        {
            List<ProvisioningJobInformation> result = new List<ProvisioningJobInformation>();

            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Get a reference to the target library
                Web web = context.Web;
                List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningJobs);

                StringBuilder sbCamlWhere = new StringBuilder();

                // Generate the CAML query filter accordingly to the requested statuses
                Boolean openCamlOr = true;
                Int32 conditionCounter = 0;
                foreach (var statusFlagName in Enum.GetNames(typeof(ProvisioningJobStatus)))
                {
                    var statusFlag = (ProvisioningJobStatus)Enum.Parse(typeof(ProvisioningJobStatus), statusFlagName);
                    if ((statusFlag & status) == statusFlag)
                    {
                        conditionCounter++;
                        if (openCamlOr)
                        {
                            // Add the first <Or /> CAML statement
                            sbCamlWhere.Insert(0, "<Or>");
                            openCamlOr = false;
                        }
                        sbCamlWhere.AppendFormat(
                            @"<Eq>
                                <FieldRef Name='PnPProvisioningJobStatus' />
                                <Value Type='Text'>" + statusFlagName + @"</Value>
                            </Eq>");

                        if (conditionCounter >= 2)
                        {
                            // Close the current <Or /> CAML statement
                            sbCamlWhere.Append("</Or>");
                            openCamlOr = true;
                        }
                    }
                }
                // Remove the first <Or> CAML statement if it is useless
                if (conditionCounter == 1)
                {
                    sbCamlWhere.Remove(0, 4);
                }

                // Add the jobType filter, if any
                if (!String.IsNullOrEmpty(jobType))
                {
                    sbCamlWhere.Insert(0, "<And>");
                    sbCamlWhere.AppendFormat(
                        @"<Eq>
                        <FieldRef Name='PnPProvisioningJobType' />
                        <Value Type='Text'>" + jobType + @"</Value>
                    </Eq>");
                    sbCamlWhere.Append("</And>");
                }

                // Add the owner filter, if any
                if (!String.IsNullOrEmpty(owner))
                {
                    Microsoft.SharePoint.Client.User ownerUser = web.EnsureUser(owner);
                    context.Load(ownerUser, u => u.Id, u => u.Email, u => u.Title);
                    context.ExecuteQueryRetry();

                    sbCamlWhere.Insert(0, "<And>");
                    sbCamlWhere.AppendFormat(
                        @"<Eq>
                        <FieldRef Name='PnPProvisioningJobOwner' />
                        <Value Type='User'>" + ownerUser.Title + @"</Value>
                    </Eq>");
                    sbCamlWhere.Append("</And>");
                }

                CamlQuery query = new CamlQuery();
                query.ViewXml =
                    @"<View>
                        <Query>
                            <Where>" + sbCamlWhere.ToString() + @"
                            </Where>
                        </Query>
                    </View>";

                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQueryRetry();

                foreach (var jobItem in items)
                {
                    result.Add(PrepareJobInformationFromSharePoint(context, jobItem, includeStream));
                }
            }
            return (result.ToArray());
        }

        public ProvisioningJob[] GetTypedProvisioningJobs<TJob>(ProvisioningJobStatus status, String owner = null)
            where TJob : ProvisioningJob
        {
            // Get the ProvisioningJobInformation array, eventually filtered by Job type
            var jobInfoList = this.GetProvisioningJobs(status,
                typeof(TJob).FullName == typeof(ProvisioningJob).FullName ? null : typeof(TJob).FullName, 
                true, owner);
            List<TJob> jobs = new List<TJob>();

            foreach (var jobInfo in jobInfoList)
            {
                jobs.Add((TJob)jobInfo.JobFile.FromJsonStream(jobInfo.Type));
            }

            return (jobs.ToArray());
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
            resultItem.ScheduledOn = (DateTime)jobItem[jobItem.ParentList.GetFieldById<FieldDateTime>(BuiltInFieldId.Created).InternalName];

            if (includeFileStream)
            {
                resultItem.JobFile = GetProvisioningJobStreamFromSharePoint(context, jobItem);
            }

            return resultItem;
        }

        private static Stream GetProvisioningJobStreamFromSharePoint(ClientContext context, ListItem jobItem)
        {
            jobItem.ParentList.RootFolder.EnsureProperty(f => f.ServerRelativeUrl);

            Microsoft.SharePoint.Client.File jobFile = jobItem.ParentList.ParentWeb.GetFileByServerRelativeUrl(
                String.Format("{0}/{1}", jobItem.ParentList.RootFolder.ServerRelativeUrl, (String)jobItem["FileLeafRef"]));
            context.Load(jobFile, jf => jf.ServerRelativeUrl);

            var jobFileStream = jobFile.OpenBinaryStream();
            context.ExecuteQueryRetry();

            MemoryStream mem = new MemoryStream();
            jobFileStream.Value.CopyTo(mem);
            mem.Position = 0;

            return (mem);
        }

        public void UpdateProvisioningJob(Guid jobId, ProvisioningJobStatus status, String errorMessage = null)
        {
            // Connect to the Infrastructural Site Collection
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Get a reference to the target library
                Web web = context.Web;
                List list = web.Lists.GetByTitle(PnPPartnerPackConstants.PnPProvisioningJobs);
                context.Load(list, l => l.RootFolder);

                CamlQuery query = new CamlQuery();
                query.ViewXml =
                    @"<View>
                        <Query>
                            <Where>
                                <Eq>
                                    <FieldRef Name='FileLeafRef' />
                                    <Value Type='Text'>" + jobId + @".job</Value>
                                </Eq>
                            </Where>
                        </Query>
                    </View>";

                ListItemCollection items = list.GetItems(query);
                context.Load(items,
                    includes => includes.IncludeWithDefaultProperties(
                        j => j[PnPPartnerPackConstants.PnPProvisioningJobStatus],
                        j => j[PnPPartnerPackConstants.PnPProvisioningJobError],
                        j => j[PnPPartnerPackConstants.PnPProvisioningJobType]),
                    includes => includes.Include(j => j.File));
                context.ExecuteQueryRetry();

                if (items.Count > 0)
                {
                    ListItem jobItem = items[0];

                    // Update the ProvisioningJob object internal status
                    ProvisioningJob job = GetProvisioningJobStreamFromSharePoint(context, jobItem)
                        .FromJsonStream((String)jobItem[PnPPartnerPackConstants.PnPProvisioningJobType]);

                    job.Status = status;
                    job.ErrorMessage = errorMessage;

                    // Update the SharePoint ListItem behind the Provisioning Job item
                    // jobItem[PnPPartnerPackConstants.ContentTypeIdField] = PnPPartnerPackConstants.PnPProvisioningJobContentTypeId;
                    jobItem[PnPPartnerPackConstants.PnPProvisioningJobStatus] = status.ToString();
                    jobItem[PnPPartnerPackConstants.PnPProvisioningJobError] = errorMessage;

                    jobItem.Update();
                    context.ExecuteQueryRetry();

                    // Update the file
                    list.RootFolder.UploadFile(jobItem.File.Name, job.ToJsonStream(), true);
                }
            }
        }
    }
}
