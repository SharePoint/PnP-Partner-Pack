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

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            IndexViewModel model = new IndexViewModel();

            model.CurrentUserPrincipalName = ClaimsPrincipal.Current.Identity.Name;

            using (var ctx = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title, w => w.Url);
                ctx.ExecuteQueryRetry();

                model.InfrastructuralSiteUrl = web.Url;
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult CreateSiteCollection()
        {
            CreateSiteCollectionViewModel model = new CreateSiteCollectionViewModel();
            return View(model);
        }

        [HttpPost]
        public ActionResult CreateSiteCollection(CreateSiteCollectionViewModel model)
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
                job.PrimarySiteCollectionAdmin = model.PrimarySiteCollectionAdmin != null &&
                    model.PrimarySiteCollectionAdmin.Length > 0 ? model.PrimarySiteCollectionAdmin[0].Login : null;
                job.SecondarySiteCollectionAdmin = model.SecondarySiteCollectionAdmin != null && 
                    model.SecondarySiteCollectionAdmin.Length > 0 ? model.SecondarySiteCollectionAdmin[0].Login : null;
                job.ProvisioningTemplateUrl = model.ProvisioningTemplateUrl;
                job.StorageMaximumLevel = model.StorageMaximumLevel;
                job.StorageWarningLevel = model.StorageWarningLevel;
                job.UserCodeMaximumLevel = model.UserCodeMaximumLevel;
                job.UserCodeWarningLevel = model.UserCodeWarningLevel;
                job.ExternalSharingEnabled = model.ExternalSharingEnabled;
                job.Title = String.Format("Provisioning of Site Collection \"{1}\" with Template \"{0}\" by {2}",
                    job.ProvisioningTemplateUrl,
                    job.RelativeUrl,
                    job.Owner);

                // TODO: Implement handling of Template Parameters
                job.TemplateParameters = null;

                model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult CreateSubSite()
        {
            CreateSubSiteViewModel model = new CreateSubSiteViewModel();
            return View(model);
        }

        [HttpPost]
        public ActionResult CreateSubSite(CreateSubSiteViewModel model)
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
                job.RelativeUrl = model.RelativeUrl;
                job.SitePolicy = model.SitePolicy;
                job.Owner = ClaimsPrincipal.Current.Identity.Name;
                job.ProvisioningTemplateUrl = model.ProvisioningTemplateUrl;
                job.InheritPermissions = model.InheritPermissions;
                job.Title = String.Format("Provisioning of Sub Site \"{1}\" with Template \"{0}\" by {2}",
                    job.ProvisioningTemplateUrl,
                    job.RelativeUrl,
                    job.Owner);

                // TODO: Implement handling of Template Parameters
                job.TemplateParameters = null;

                model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult SaveSiteAsTemplate(String spHostUrl)
        {
            SaveTemplateViewModel model = new SaveTemplateViewModel();
            model.SourceSiteUrl = spHostUrl;
            return View(model);
        }

        [HttpPost]
        public ActionResult SaveSiteAsTemplate(SaveTemplateViewModel model)
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
                job.Title = String.Format("Saving of Template \"{0}\" from Site \"{1}\" by {2}",
                    job.FileName,
                    job.SourceSiteUrl,
                    job.Owner);

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
                     select new SiteCollectionItem {
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
            if (ModelState.IsValid)
            {
                PnPPartnerPackUtilities.EnablePartnerPackOnSite("https://piasysdev.sharepoint.com/sites/PnPProvisioningTarget03/");
            }

            return View("Index");
        }

        [HttpGet]
        public ActionResult MyProvisionedSites()
        {
            return View("MyProvisionedSites");
        }

        [HttpPost]
        public ActionResult GetPeoplePickerData()
        {
            return Json(PeoplePickerHelper.GetPeoplePickerSearchData());
        }

        [HttpPost]
        public ActionResult GetSiteCollectionSettings(String siteCollectionUri)
        {
            return Json(PnPPartnerPackUtilities.GetSiteCollectionSettings(siteCollectionUri));
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

            return Json(PnPPartnerPackUtilities.GetSiteCollectionSettings(siteCollectionUri));
        }
    }
}