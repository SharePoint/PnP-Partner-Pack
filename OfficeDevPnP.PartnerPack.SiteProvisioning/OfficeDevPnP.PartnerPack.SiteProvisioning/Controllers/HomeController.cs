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
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title, w => w.Url);
                ctx.ExecuteQueryRetry();

                model.InfrastructuralSiteUrl = web.Url;
            }

            var currentUser = GetCurrentUser();
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
            model.Scope = TemplateScope.Site;
            model.ParentSiteUrl = String.Empty;
            model.PartnerPackExtensionsEnabled = true;
            model.ResponsiveDesignEnabled = true;
            return View("CreateSite", model);
        }

        [HttpPost]
        public ActionResult CreateSiteCollection(CreateSiteCollectionViewModel model)
        {
            switch (model.Step)
            {
                case CreateSiteStep.SiteInformation:
                    ModelState.Clear();
                    if (String.IsNullOrEmpty(model.Title))
                    {
                        // Set initial value for PnP Partner Pack Extensions Enabled
                        model.PartnerPackExtensionsEnabled = true;
                        model.ResponsiveDesignEnabled = true;
                    }
                    break;
                case CreateSiteStep.TemplateParameters:
                    if (!ModelState.IsValid)
                    {
                        model.Step = CreateSiteStep.SiteInformation;
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(model.ProvisioningTemplateUrl) &&
                            model.ProvisioningTemplateUrl.IndexOf(PnPPartnerPackConstants.PnPProvisioningTemplates) > 0)
                        {
                            String templateSiteUrl = model.ProvisioningTemplateUrl.Substring(0, model.ProvisioningTemplateUrl.IndexOf(PnPPartnerPackConstants.PnPProvisioningTemplates));
                            String templateFileName = model.ProvisioningTemplateUrl.Substring(model.ProvisioningTemplateUrl.IndexOf(PnPPartnerPackConstants.PnPProvisioningTemplates) + PnPPartnerPackConstants.PnPProvisioningTemplates.Length + 1);
                            String templateFolder = String.Empty;

                            if (templateFileName.IndexOf("/") > 0)
                            {
                                templateFolder = templateFileName.Substring(0, templateFileName.LastIndexOf("/") - 1);
                                templateFileName = templateFileName.Substring(templateFolder.Length + 1);
                            }
                            model.TemplateParameters = PnPPartnerPackUtilities.GetProvisioningTemplateParameters(
                                    templateSiteUrl,
                                    templateFolder,
                                    templateFileName);
                        }
                    }
                    break;
                case CreateSiteStep.SiteCreated:
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
                        job.PrimarySiteCollectionAdmin = model.PrimarySiteCollectionAdmin != null &&
                            model.PrimarySiteCollectionAdmin.Length > 0 ? model.PrimarySiteCollectionAdmin[0].Email : null;
                        job.SecondarySiteCollectionAdmin = model.SecondarySiteCollectionAdmin != null &&
                            model.SecondarySiteCollectionAdmin.Length > 0 ? model.SecondarySiteCollectionAdmin[0].Email : null;
                        job.ProvisioningTemplateUrl = model.ProvisioningTemplateUrl;
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
                    break;
                default:
                    break;
            }

            return PartialView(model.Step.ToString(), model);
        }

        [HttpGet]
        public ActionResult CreateSubSite()
        {
            CreateSubSiteViewModel model = new CreateSubSiteViewModel();
            model.Scope = TemplateScope.Web;
            model.ParentSiteUrl = HttpContext.Request["SPHostUrl"];

            return View("CreateSite", model);
        }

        [HttpPost]
        public ActionResult CreateSubSite(CreateSubSiteViewModel model)
        {
            switch (model.Step)
            {
                case CreateSiteStep.SiteInformation:
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
                    break;
                case CreateSiteStep.TemplateParameters:
                    if (!ModelState.IsValid)
                    {
                        model.Step = CreateSiteStep.SiteInformation;
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(model.ProvisioningTemplateUrl) &&
                            model.ProvisioningTemplateUrl.IndexOf(PnPPartnerPackConstants.PnPProvisioningTemplates) > 0)
                        {
                            String templateSiteUrl = model.ProvisioningTemplateUrl.Substring(0, model.ProvisioningTemplateUrl.IndexOf(PnPPartnerPackConstants.PnPProvisioningTemplates));
                            String templateFileName = model.ProvisioningTemplateUrl.Substring(model.ProvisioningTemplateUrl.IndexOf(PnPPartnerPackConstants.PnPProvisioningTemplates) + PnPPartnerPackConstants.PnPProvisioningTemplates.Length + 1);
                            String templateFolder = String.Empty;

                            if (templateFileName.IndexOf("/") > 0)
                            {
                                templateFolder = templateFileName.Substring(0, templateFileName.LastIndexOf("/") - 1);
                                templateFileName = templateFileName.Substring(templateFolder.Length + 1);
                            }
                            model.TemplateParameters = PnPPartnerPackUtilities.GetProvisioningTemplateParameters(
                                    templateSiteUrl,
                                    templateFolder,
                                    templateFileName);
                        }
                    }
                    break;
                case CreateSiteStep.SiteCreated:
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
                        job.ProvisioningTemplateUrl = model.ProvisioningTemplateUrl;
                        job.InheritPermissions = model.InheritPermissions;
                        job.Title = String.Format("Provisioning of Sub Site \"{1}\" with Template \"{0}\" by {2}",
                            job.ProvisioningTemplateUrl,
                            job.RelativeUrl,
                            job.Owner);

                        job.TemplateParameters = model.TemplateParameters;

                        model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
                    }
                    break;
                default:
                    break;
            }

            return PartialView(model.Step.ToString(), model);
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
                        job.Scope = TemplateScope.Site;
                        storageLocationUrl = rootWeb.Url;
                    }
                    else
                    {
                        // Otherwise we are in a Sub Site of the Site Collection
                        job.Scope = TemplateScope.Web;
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
        public ActionResult GetPeoplePickerData()
        {
            return Json(PeoplePickerHelper.GetPeoplePickerSearchData());
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

        public ActionResult GetPersonaPhoto(String upn, Int32 width = 0, Int32 height = 0)
        {
            Stream result = null;
            String contentType = "image/png";

            var sourceStream = GetUserPhoto(upn);

            if (sourceStream != null && width != 0 && height != 0)
            {
                Image sourceImage = Image.FromStream(sourceStream);
                Image resultImage = ScaleImage(sourceImage, width, height);

                result = new MemoryStream();
                resultImage.Save(result, ImageFormat.Png);
                result.Position = 0;
            }
            else
            {
                result = sourceStream;
            }

            if (result != null)
            {
                return base.File(result, contentType);
            }
            else
            {
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.NoContent);
            }
        }

        /// <summary>
        /// This method retrieves the current user from Azure AD
        /// </summary>
        /// <returns>The user retrieved from Azure AD</returns>
        public static LightGraphUser GetCurrentUser()
        {
            String jsonResponse = MicrosoftGraphHelper.MakeGetRequestForString(
                String.Format("{0}me",
                    MicrosoftGraphHelper.MicrosoftGraphV1BaseUri));

            if (jsonResponse != null)
            {
                var user = JsonConvert.DeserializeObject<LightGraphUser>(jsonResponse);
                return (user);
            }
            else
            {
                return (null);
            }
        }

        /// <summary>
        /// This method retrieves the photo of a single user from Azure AD
        /// </summary>
        /// <param name="upn">The UPN of the user</param>
        /// <returns>The user's photo retrieved from Azure AD</returns>
        private static Stream GetUserPhoto(String upn)
        {
            String contentType = "image/png";

            var result = MicrosoftGraphHelper.MakeGetRequestForStream(
                String.Format("{0}users/{1}/photo/$value",
                    MicrosoftGraphHelper.MicrosoftGraphV1BaseUri, upn),
                contentType);

            return (result);
        }

        private Image ScaleImage(Image image, int maxWidth, int maxHeight)
        {
            var ratioX = (double)maxWidth / image.Width;
            var ratioY = (double)maxHeight / image.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(image.Width * ratio);
            var newHeight = (int)(image.Height * ratio);

            var newImage = new Bitmap(newWidth, newHeight);

            using (var graphics = Graphics.FromImage(newImage))
                graphics.DrawImage(image, 0, 0, newWidth, newHeight);

            return newImage;
        }
    }
}