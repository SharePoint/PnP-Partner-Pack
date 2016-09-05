using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.PartnerPack.Infrastructure;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using System.Xml.Serialization;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Controllers
{
    public class GovernanceController : Controller
    {
        [HttpGet]
        public ActionResult Branding()
        {
            BrandingViewModel model = getCurrentBrandingSettings();
            return View(model);
        }

        [HttpPost]
        public ActionResult Branding(BrandingViewModel model)
        {
            AntiForgery.Validate();
            if (ModelState.IsValid)
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
                var job = new ApplyBrandingJob
                {
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
            }

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

        [HttpGet]
        public ActionResult UpdateTemplates()
        {
            UpdateTemplatesViewModel model = new UpdateTemplatesViewModel();
            return View(model);
        }

        [HttpPost]
        public ActionResult UpdateTemplates(UpdateTemplatesViewModel model)
        {
            AntiForgery.Validate();
            if (ModelState.IsValid)
            {
                // Create the asynchronous job
                var job = new UpdateTemplatesJob
                {
                    Owner = ClaimsPrincipal.Current.Identity.Name,
                    Title = "Tenant Wide Update Templates",
                };

                // Enqueue the job for execution
                model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult SiteCollectionsBatch()
        {
            var model = new SiteCollectionsBatchViewModel();
            return View(model);
        }

        [HttpPost]
        public ActionResult SiteCollectionsBatch(SiteCollectionsBatchViewModel model, HttpPostedFileBase batchFile)
        {
            switch (model.Step)
            {
                case BatchStep.BatchStartup:
                    ModelState.Clear();
                    break;
                case BatchStep.BatchFileUploaded:
                    if (!ModelState.IsValid)
                    {
                        model.Step = BatchStep.BatchStartup;
                    }
                    else
                    {
                        try
                        {
                            XmlSerializer xs = new XmlSerializer(typeof(batches));
                            model.Sites = (batches)xs.Deserialize(batchFile.InputStream);
                            model.SitesJson = JsonConvert.SerializeObject(model.Sites);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Invalid batch file!", ex);
                        }
                    }
                    break;
                case BatchStep.BatchScheduled:
                    if (!String.IsNullOrEmpty(model.SitesJson))
                    {
                        // Create the asynchronous job
                        var job = new SiteCollectionsBatchJob
                        {
                            Owner = ClaimsPrincipal.Current.Identity.Name,
                            Title = "Site Collections Batch",
                            BatchSites = model.SitesJson,
                        };

                        // Enqueue the job for execution
                        model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
                    }
                    break;
                default:
                    break;
            }

            return PartialView(model.Step.ToString(), model);
        }
    }
}