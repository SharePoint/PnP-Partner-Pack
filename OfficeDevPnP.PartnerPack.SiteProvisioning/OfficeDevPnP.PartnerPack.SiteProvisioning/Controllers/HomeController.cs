using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Components;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Models;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Web;
using System.Web.Mvc;

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

        public ActionResult CreateSiteCollection()
        {
            CreateSiteCollectionViewModel model = new CreateSiteCollectionViewModel();
            return View(model);
        }

        public ActionResult CreateSubSite()
        {
            CreateSubSiteViewModel model = new CreateSubSiteViewModel();
            return View(model);
        }

        public ActionResult MyProvisionedSites()
        {
            return View("MyProvisionedSites");
        }

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
            PnPPartnerPackUtilities.EnablePartnerPackOnSite("https://piasysdev.sharepoint.com/sites/PnPProvisioningTarget03/");
            return View("Index");
        }

        public ActionResult SaveSiteAsTemplate()
        {
            SaveTemplateViewModel model = new SaveTemplateViewModel();
            return View(model);
        }

        public ActionResult ApplyProvisioningTemplate()
        {
            return View("ApplyProvisioningTemplate");
        }

        [HttpPost]
        public ActionResult GetPeoplePickerData()
        {
            return Json(PeoplePickerHelper.GetPeoplePickerSearchData());
        }

        private PeoplePickerUser GetApplicant(string loginName)
        {
            using (var ctx = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                return GetApplicant(ctx.Web.EnsureUser(loginName));
            }
        }

        private PeoplePickerUser GetApplicant(Microsoft.SharePoint.Client.User user)
        {
            using (var ctx = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                ctx.Load(user, u => u.LoginName, u => u.Title, u => u.Email);
                ctx.ExecuteQuery();

                return new PeoplePickerUser
                {
                    Login = user.LoginName,
                    Email = user.Email,
                    Name = user.Title,
                };
            }
        }
    }
}