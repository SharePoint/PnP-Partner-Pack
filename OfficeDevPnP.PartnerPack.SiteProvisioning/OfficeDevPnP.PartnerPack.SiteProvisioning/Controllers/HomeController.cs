using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Components;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Models;
using OfficeDevPnP.PartnerPack.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Web;
using System.Web.Mvc;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System.Web.Helpers;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Net.Http;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System.Net.Http.Headers;
using System.Net;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            IndexViewModel model = new IndexViewModel();

            using (var ctx = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                // Track usage of the PnP Partner Pack
                ctx.ClientTag = "SPDev:PartnerPack";

                Web web = ctx.Web;
                ctx.Load(web, w => w.Title, w => w.Url);
                ctx.ExecuteQuery();

                model.InfrastructuralSiteUrl = web.Url;
            }

            var currentUser = UserUtility.GetCurrentUser();
            if (currentUser != null)
            {
                model.CurrentUserPrincipalName = currentUser.UserPrincipalName;
            }
            else
            {
                model.CurrentUserPrincipalName = ClaimsPrincipal.Current.Identity.Name;
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult CreateSiteCollection()
        {
            CreateSiteCollectionViewModel model = new CreateSiteCollectionViewModel();
            model.Scope = TargetScope.Site;
            model.ParentSiteUrl = String.Empty;
            model.PartnerPackExtensionsEnabled = true;
            model.ResponsiveDesignEnabled = true;
            return View("CreateSite", model);
        }

        [HttpPost]
        public ActionResult CreateSiteCollection(CreateSiteCollectionViewModel model)
        {

            if (model.Step == CreateSiteStep.SiteInformation)
            {
                ModelState.Clear();
                if (String.IsNullOrEmpty(model.Title))
                {
                    // Set initial value for PnP Partner Pack Extensions Enabled
                    model.PartnerPackExtensionsEnabled = true;
                    model.ResponsiveDesignEnabled = true;
                }
            }
            if (model.Step == CreateSiteStep.TemplateParameters)
            {
                if (!ModelState.IsValid)
                {
                    model.Step = CreateSiteStep.SiteInformation;
                }
                else
                {
                    if (!String.IsNullOrEmpty(model.ProvisioningTemplateUrl) &&
                        !String.IsNullOrEmpty(model.TemplatesProviderTypeName))
                    {
                        var templatesProvider = PnPPartnerPackSettings.TemplatesProviders[model.TemplatesProviderTypeName];
                        if (templatesProvider != null)
                        {
                            var template = templatesProvider.GetProvisioningTemplate(model.ProvisioningTemplateUrl);
                            model.TemplateParameters = template.Parameters;
                        }
                    }

                    if (model.TemplateParameters == null || model.TemplateParameters.Count == 0)
                    {
                        model.Step = CreateSiteStep.SiteCreated;
                    }
                }
            }
            if (model.Step == CreateSiteStep.SiteCreated)
            {
                AntiForgery.Validate();
                if (ModelState.IsValid)
                {
                    // Prepare the Job to provision the Site Collection
                    SiteCollectionProvisioningJob job = new SiteCollectionProvisioningJob();

                    // Prepare all the other information about the Provisioning Job
                    job.SiteTitle = model.Title;
                    job.Description = model.Description;
                    job.Language = model.Language;
                    job.TimeZone = model.TimeZone;
                    job.RelativeUrl = String.Format("/{0}/{1}", model.ManagedPath, model.RelativeUrl);
                    job.SitePolicy = model.SitePolicy;
                    job.Owner = ClaimsPrincipal.Current.Identity.Name;
                    job.ApplyTenantBranding = model.ApplyTenantBranding;

                    job.PrimarySiteCollectionAdmin = model.PrimarySiteCollectionAdmin != null &&
                        model.PrimarySiteCollectionAdmin.Principals.Count > 0 ?
                            (!String.IsNullOrEmpty(model.PrimarySiteCollectionAdmin.Principals[0].Mail) ?
                            model.PrimarySiteCollectionAdmin.Principals[0].Mail :
                            null) : null;
                    job.SecondarySiteCollectionAdmin = model.SecondarySiteCollectionAdmin != null &&
                        model.SecondarySiteCollectionAdmin.Principals.Count > 0 ?
                            (!String.IsNullOrEmpty(model.SecondarySiteCollectionAdmin.Principals[0].Mail) ?
                            model.SecondarySiteCollectionAdmin.Principals[0].Mail :
                            null) : null;

                    job.ProvisioningTemplateUrl = model.ProvisioningTemplateUrl;
                    job.TemplatesProviderTypeName = model.TemplatesProviderTypeName;
                    job.StorageMaximumLevel = model.StorageMaximumLevel;
                    job.StorageWarningLevel = model.StorageWarningLevel;
                    job.UserCodeMaximumLevel = model.UserCodeMaximumLevel;
                    job.UserCodeWarningLevel = model.UserCodeWarningLevel;
                    job.ExternalSharingEnabled = model.ExternalSharingEnabled;
                    job.ResponsiveDesignEnabled = model.ResponsiveDesignEnabled;
                    job.PartnerPackExtensionsEnabled = model.PartnerPackExtensionsEnabled;
                    job.Title = String.Format("Provisioning of Site Collection \"{1}\" with Template \"{0}\" by {2}",
                        job.ProvisioningTemplateUrl,
                        job.RelativeUrl,
                        job.Owner);

                    job.TemplateParameters = model.TemplateParameters;

                    model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
                }
            }

            return PartialView(model.Step.ToString(), model);
        }

        [HttpGet]
        public ActionResult CreateSubSite()
        {
            CreateSubSiteViewModel model = new CreateSubSiteViewModel();
            model.Scope = TargetScope.Web;
            model.ParentSiteUrl = HttpContext.Request["SPHostUrl"];

            PnPPartnerPackSettings.ParentSiteUrl = model.ParentSiteUrl;

            return View("CreateSite", model);
        }

        [HttpPost]
        public ActionResult CreateSubSite(CreateSubSiteViewModel model)
        {
            PnPPartnerPackSettings.ParentSiteUrl = model.ParentSiteUrl;

            if (model.Step == CreateSiteStep.SiteInformation)
            {
                ModelState.Clear();

                // If it is the first time that we are here
                if (String.IsNullOrEmpty(model.Title))
                {
                    model.InheritPermissions = true;
                    using (var ctx = PnPPartnerPackContextProvider.GetAppOnlyClientContext(model.ParentSiteUrl))
                    {
                        Web web = ctx.Web;
                        ctx.Load(web, w => w.Language, w => w.RegionalSettings.TimeZone);
                        ctx.ExecuteQueryRetry();

                        model.Language = (Int32)web.Language;
                        model.TimeZone = web.RegionalSettings.TimeZone.Id;
                    }
                }
            }
            if (model.Step == CreateSiteStep.TemplateParameters)
            {
                if (!ModelState.IsValid)
                {
                    model.Step = CreateSiteStep.SiteInformation;
                }
                else
                {
                    if (!String.IsNullOrEmpty(model.ProvisioningTemplateUrl) &&
                        !String.IsNullOrEmpty(model.TemplatesProviderTypeName))
                    {
                        var templatesProvider = PnPPartnerPackSettings.TemplatesProviders[model.TemplatesProviderTypeName];
                        if (templatesProvider != null)
                        {
                            var template = templatesProvider.GetProvisioningTemplate(model.ProvisioningTemplateUrl);
                            model.TemplateParameters = template.Parameters;
                        }

                        if (model.TemplateParameters == null || model.TemplateParameters.Count == 0)
                        {
                            model.Step = CreateSiteStep.SiteCreated;
                        }
                    }
                }
            }
            if (model.Step == CreateSiteStep.SiteCreated)
            {
                AntiForgery.Validate();
                if (ModelState.IsValid)
                {
                    // Prepare the Job to provision the Sub Site 
                    SubSiteProvisioningJob job = new SubSiteProvisioningJob();

                    // Prepare all the other information about the Provisioning Job
                    job.SiteTitle = model.Title;
                    job.Description = model.Description;
                    job.Language = model.Language;
                    job.TimeZone = model.TimeZone;
                    job.ParentSiteUrl = model.ParentSiteUrl;
                    job.RelativeUrl = model.RelativeUrl;
                    job.SitePolicy = model.SitePolicy;
                    job.Owner = ClaimsPrincipal.Current.Identity.Name;
                    job.ApplyTenantBranding = model.ApplyTenantBranding;

                    job.ProvisioningTemplateUrl = model.ProvisioningTemplateUrl;
                    job.TemplatesProviderTypeName = model.TemplatesProviderTypeName;
                    job.InheritPermissions = model.InheritPermissions;
                    job.Title = String.Format("Provisioning of Sub Site \"{1}\" with Template \"{0}\" by {2}",
                        job.ProvisioningTemplateUrl,
                        job.RelativeUrl,
                        job.Owner);

                    job.TemplateParameters = model.TemplateParameters;

                    model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
                }
            }

            return PartialView(model.Step.ToString(), model);
        }

        [HttpPost]
        public ActionResult SearchTemplates(OfficeDevPnP.PartnerPack.Infrastructure.TargetScope scope, String parentSiteUrl, String templatesProvider, String searchText)
        {
            PnPPartnerPackSettings.ParentSiteUrl = parentSiteUrl;

            SearchTemplatesViewModel model = new SearchTemplatesViewModel();

            if (!PnPPartnerPackSettings.TemplatesProviders.ContainsKey(templatesProvider))
            {
                throw new Exception("Invalid templates provider key!");
            }

            model.SearchResults = PnPPartnerPackSettings.TemplatesProviders[templatesProvider]
                .SearchProvisioningTemplates(searchText, TargetPlatform.SharePointOnline, scope);

            //List<ProvisioningTemplateInformation> result = new List<ProvisioningTemplateInformation>();

            //var globalTemplates = ProvisioningRepositoryFactory.Current.GetGlobalProvisioningTemplates(scope);
            //result.AddRange(globalTemplates);

            //if (scope != TargetScope.Site)
            //{
            //    var localTemplates = ProvisioningRepositoryFactory.Current.GetLocalProvisioningTemplates(parentSiteUrl, scope);
            //    result.AddRange(localTemplates);
            //}

            //model.SearchResults = result.ToArray();

            return PartialView(model);
        }

        [HttpGet]
        public ActionResult SaveSiteAsTemplate(String spHostUrl)
        {
            SaveTemplateViewModel model = new SaveTemplateViewModel();
            model.SourceSiteUrl = spHostUrl;
            model.IncludeAllTermGroups = false;
            model.IncludeSiteCollectionTermGroup = false;
            return View(model);
        }

        [HttpPost]
        public ActionResult SaveSiteAsTemplate(SaveTemplateViewModel model, HttpPostedFileBase templateImageFile)
        {
            AntiForgery.Validate();
            if (ModelState.IsValid)
            {
                // Prepare the Job to store the Provisioning Template
                GetProvisioningTemplateJob job = new GetProvisioningTemplateJob();

                // Store the local location for the Provisioning Template, if any
                String storageLocationUrl = null;

                // Determine the Scope of the Provisioning Template
                using (var ctx = PnPPartnerPackContextProvider.GetAppOnlyClientContext(model.SourceSiteUrl))
                {
                    Web web = ctx.Web;
                    Web rootWeb = ctx.Site.RootWeb;
                    ctx.Load(web, w => w.Id);
                    ctx.Load(rootWeb, w => w.Url, w => w.Id);
                    ctx.ExecuteQueryRetry();

                    if (web.Id == rootWeb.Id)
                    {
                        // We are in the Root Site of the Site Collection
                        job.Scope = TargetScope.Site;
                        storageLocationUrl = rootWeb.Url;
                    }
                    else
                    {
                        // Otherwise we are in a Sub Site of the Site Collection
                        job.Scope = TargetScope.Web;
                    }
                }

                // Prepare all the other information about the Provisioning Job
                job.Owner = ClaimsPrincipal.Current.Identity.Name;
                job.FileName = model.FileName;
                job.IncludeAllTermGroups = model.IncludeAllTermGroups;
                job.IncludeSearchConfiguration = model.IncludeSearchConfiguration;
                job.IncludeSiteCollectionTermGroup = model.IncludeSiteCollectionTermGroup;
                job.IncludeSiteGroups = model.IncludeSiteGroups;
                job.PersistComposedLookFiles = model.PersistComposedLookFiles;
                job.SourceSiteUrl = model.SourceSiteUrl;
                job.Title = model.Title;
                job.Description = model.Description;
                job.Location = (ProvisioningTemplateLocation)Enum.Parse(typeof(ProvisioningTemplateLocation), model.Location, true);
                job.StorageSiteLocationUrl = storageLocationUrl;
                if (templateImageFile != null && templateImageFile.ContentLength > 0)
                {
                    job.TemplateImageFile = templateImageFile.InputStream.FixedSizeImageStream(320, 180).ToByteArray();
                    job.TemplateImageFileName = templateImageFile.FileName;
                }

                model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult UpdateSiteTemplate(String spHostUrl)
        {
            UpdateSiteTemplateViewModel model = new UpdateSiteTemplateViewModel();
            model.TargetSiteUrl = spHostUrl;
            return View(model);
        }

        [HttpPost]
        public ActionResult UpdateSiteTemplate(UpdateSiteTemplateViewModel model)
        {
            AntiForgery.Validate();
            if (ModelState.IsValid)
            {
                // Prepare the Job to update the Provisioning Template
                RefreshSingleSiteJob job = new RefreshSingleSiteJob();
                job.TargetSiteUrl = model.TargetSiteUrl;

                // Prepare all the other information about the Provisioning Job
                job.Owner = ClaimsPrincipal.Current.Identity.Name;
                job.Title = $"Update Provisioning Template for {model.TargetSiteUrl}";

                model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult ApplyProvisioningTemplate()
        {
            ApplyProvisioningTemplateViewModel model = new ApplyProvisioningTemplateViewModel();
            return View(model);
        }

        [HttpPost]
        public ActionResult ApplyProvisioningTemplate(ApplyProvisioningTemplateViewModel model)
        {
            AntiForgery.Validate();
            if (ModelState.IsValid)
            {
                // Prepare the Job to apply the Provisioning Template
                ApplyProvisioningTemplateJob job = new ApplyProvisioningTemplateJob();

                // Prepare all the other information about the Provisioning Job
                job.Owner = ClaimsPrincipal.Current.Identity.Name;
                job.ProvisioningTemplateUrl = model.ProvisioningTemplateUrl;
                job.TargetSiteUrl = model.RelativeUrl;
                job.Title = String.Format("Application of Template \"{0}\" to Site \"{1}\" by {2}",
                    job.ProvisioningTemplateUrl,
                    job.TargetSiteUrl,
                    job.Owner);

                model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult Settings()
        {
            SettingsViewModel model = new SettingsViewModel();
            
            using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContext())
            {
                var tenant = new Tenant(adminContext);

                // TODO: Here we could add paging capabilities
                var siteCollections = tenant.GetSiteProperties(0, true);
                adminContext.Load(siteCollections);
                adminContext.ExecuteQueryRetry();

                model.SiteCollections =
                    (from site in siteCollections
                     select new SiteCollectionSettings {
                         Title = site.Title,
                         Url = site.Url,
                         PnPPartnerPackEnabled = false, // PnPPartnerPackUtilities.IsPartnerPackEnabledOnSite(site.Url),
                     }).ToArray();
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult Settings(SettingsViewModel model)
        {
            AntiForgery.Validate();

            return View("Index");
        }

        [HttpGet]
        public ActionResult MyProvisionedSites()
        {
            return View("MyProvisionedSites");
        }

        [HttpPost]
        public ActionResult GetMyProvisionedSitesList()
        {
            MyProvisionedSitesViewModel model = new MyProvisionedSitesViewModel();

            // Get all the jobs related to Site Collections provisioning, enqueued by the current user
            model.PersonalJobs = ProvisioningRepositoryFactory.Current.GetTypedProvisioningJobs<SiteCollectionProvisioningJob>(
                ProvisioningJobStatus.Pending | ProvisioningJobStatus.Running |
                ProvisioningJobStatus.Provisioned | ProvisioningJobStatus.Failed |
                ProvisioningJobStatus.Cancelled,
                ClaimsPrincipal.Current.Identity.Name);

            return PartialView("MyProvisionedSitesGrid", model);
        }

        [HttpPost]
        public ActionResult GetSiteCollectionSettings(String siteCollectionUri)
        {
            return PartialView(PnPPartnerPackUtilities.GetSiteCollectionSettings(siteCollectionUri));
        }

        [HttpPost]
        public ActionResult ToggleSiteCollectionSettings(String siteCollectionUri, Boolean toggleAction)
        {
            if (toggleAction)
            {
                PnPPartnerPackUtilities.EnablePartnerPackOnSite(siteCollectionUri);
            }
            else
            {
                PnPPartnerPackUtilities.DisablePartnerPackOnSite(siteCollectionUri);
            }

            return PartialView("GetSiteCollectionSettings", PnPPartnerPackUtilities.GetSiteCollectionSettings(siteCollectionUri));
        }

        [AllowAnonymous]
        public ActionResult Error(string message)
        {
            throw new Exception(message);
        }

        [HttpGet]
        public ActionResult GetTemplateImagePreviewFromPnP(String imagePreviewUri)
        {
            if (String.IsNullOrEmpty(imagePreviewUri))
            {
                throw new ArgumentNullException("imagePreviewUri");
            }

            // Recover the original protocol moniker
            var sourceUrl = imagePreviewUri.Replace("pnps://", "https://").Replace("pnp://", "http://");
            var imagePreviewFile = sourceUrl.Substring(sourceUrl.LastIndexOf("/") + 1);
            var pnpFileUrl = sourceUrl.Substring(0, sourceUrl.LastIndexOf("/"));
            var pnpFileName = pnpFileUrl.Substring(pnpFileUrl.LastIndexOf("/") + 1);
            var sourceSiteUrl = PnPPartnerPackUtilities.GetSiteCollectionRootUrl(pnpFileUrl);
            var sourceSiteFolder = pnpFileUrl.Substring(sourceSiteUrl.Length + 1, pnpFileUrl.LastIndexOf("/") - sourceSiteUrl.Length - 1);

            using (var repositoryContext = PnPPartnerPackContextProvider.GetAppOnlyClientContext(sourceSiteUrl))
            {
                var repositoryWeb = repositoryContext.Web;
                repositoryWeb.EnsureProperty(w => w.Url);

                XMLTemplateProvider provider = new XMLOpenXMLTemplateProvider(pnpFileName,
                    new SharePointConnector(repositoryContext, repositoryWeb.Url, sourceSiteFolder));

                var imageFileStream = provider.Connector.GetFileStream(imagePreviewFile);

                return (base.File(imageFileStream, "image/png"));
            }
        }
    }
}
