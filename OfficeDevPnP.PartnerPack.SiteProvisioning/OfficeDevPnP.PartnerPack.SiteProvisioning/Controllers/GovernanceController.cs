using Newtonsoft.Json;
using OfficeDevPnP.PartnerPack.Infrastructure;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Controllers
{
    public class GovernanceController : Controller
    {
        // GET: Governance/ExternalSharing
        public ActionResult ExternalSharing()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Branding()
        {
            BrandingViewModel model = getCurrentBrandingSettings();
            return View(model);
        }

        [HttpPost]
        public ActionResult Branding(BrandingViewModel model)
        {
            // Get the original branding settings
            BrandingViewModel original = getCurrentBrandingSettings();

            // If something changed in the branding settings compared with previous settings
            if (model.LogoImageUrl != original.LogoImageUrl ||
                model.BackgroundImageUrl != original.BackgroundImageUrl ||
                model.ColorFileUrl != original.ColorFileUrl ||
                model.FontFileUrl != original.FontFileUrl ||
                model.CSSOverrideUrl != original.CSSOverrideUrl ||
                model.UICustomActionsUrl != original.UICustomActionsUrl)
            {
                // Update the last date and time of update
                model.UpdatedOn = DateTime.Now;

                // Get the JSON representation of the settings
                var jsonBranding = JsonConvert.SerializeObject(model);

                // Store the new tenant-wide settings in the Infrastructural Site Collection
                PnPPartnerPackUtilities.SetPropertyBagValueToInfrastructure(
                    PnPPartnerPackConstants.PropertyBag_Branding, jsonBranding);
            }

            // Create the asynchronous job
            var job = new ApplyBrandingJob {
                LogoImageUrl = model.LogoImageUrl,
                BackgroundImageUrl = model.BackgroundImageUrl,
                ColorFileUrl = model.ColorFileUrl,
                FontFileUrl = model.FontFileUrl,
                CSSOverrideUrl = model.CSSOverrideUrl,
                UICustomActionsUrl = model.UICustomActionsUrl,
                Owner = ClaimsPrincipal.Current.Identity.Name,
                Title = "Tenant Wide Branding",                
            };

            // Enqueue the job for execution
            model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);

            return View(model);
        }

        /// <summary>
        /// Private method for reading branding settings from the Infrastructural Site Collection
        /// </summary>
        /// <returns></returns>
        private static BrandingViewModel getCurrentBrandingSettings()
        {
            // Get the current settings from the Infrastructural Site Collection
            var jsonBrandingSettings = PnPPartnerPackUtilities.GetPropertyBagValueFromInfrastructure(
                PnPPartnerPackConstants.PropertyBag_Branding);

            // Read the current branding settings, if any
            var model = jsonBrandingSettings != null ?
                JsonConvert.DeserializeObject<BrandingViewModel>(jsonBrandingSettings) :
                new BrandingViewModel();
            return model;
        }

        // GET: Governance/UpdateTemplates
        public ActionResult UpdateTemplates()
        {
            return View();
        }

        // GET: Governance/SiteCollectionsBatch
        public ActionResult SiteCollectionsBatch()
        {
            return View();
        }
    }
}