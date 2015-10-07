using Microsoft.SharePoint.Client;
using OfficeDevPnP.PartnerPack.SiteProvisioningWeb.Components;
using OfficeDevPnP.PartnerPack.SiteProvisioningWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace OfficeDevPnP.PartnerPack.SiteProvisioningWeb.Controllers
{
    [SharePointContextFilter]
    public class HomeController : Controller
    {
        private ClientContext retrieveClientContext()
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            return (spContext.CreateUserClientContextForSPHost());
        }

        public ActionResult Index()
        {
            return View();
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

        public ActionResult AddRemoveExtensions()
        {
            String scriptName = "PnPPartnerPackOverrides";

            using (var ctx = retrieveClientContext())
            {
                if (!JSLinkUtility.ExistsJsLink(scriptName, ctx, ctx.Web))
                {
                    JSLinkUtility.AddJsLink(scriptName, "Scripts/PnP-Partner-Pack-Overrides.js", this.Request, ctx, ctx.Web);
                }
                else
                {
                    JSLinkUtility.DeleteJsLink(scriptName, ctx, ctx.Web);
                }
            }


            return View("Index");
        }
        public ActionResult SaveSiteTemplate()
        {
            SaveTemplateViewModel model = new SaveTemplateViewModel();
            return View(model);
        }

        public ActionResult ApplyProvisioningTemplate()
        {
            return View("Index");
        }

        [HttpPost]
        public ActionResult GetPeoplePickerData()
        {
            return Json(PeoplePickerHelper.GetPeoplePickerSearchData());
        }

        private PeoplePickerUser GetApplicant(string loginName)
        {
            using (var ctx = retrieveClientContext())
            {
                return GetApplicant(ctx.Web.EnsureUser(loginName));
            }
        }

        private PeoplePickerUser GetApplicant(Microsoft.SharePoint.Client.User user)
        {
            using (var ctx = retrieveClientContext())
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