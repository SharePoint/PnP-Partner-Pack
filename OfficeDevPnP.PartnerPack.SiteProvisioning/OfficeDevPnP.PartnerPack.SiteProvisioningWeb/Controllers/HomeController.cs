using Microsoft.SharePoint.Client;
using OfficeDevPnP.PartnerPack.SiteProvisioningWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
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
    }
}