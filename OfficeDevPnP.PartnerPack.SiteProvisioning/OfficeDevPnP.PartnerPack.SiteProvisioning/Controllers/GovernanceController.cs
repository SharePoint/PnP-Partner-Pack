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
            BrandingSettings original = PnPPartnerPackUtilities.GetTenantBrandingSettings();

            BrandingViewModel model = new BrandingViewModel {
                LogoImageUrl = original.LogoImageUrl,
                BackgroundImageUrl = original.BackgroundImageUrl,
                ColorFileUrl = original.ColorFileUrl,
                FontFileUrl = original.FontFileUrl,
                CSSOverrideUrl = original.CSSOverrideUrl,
                UICustomActionsUrl = original.UICustomActionsUrl,
            };

            return View(model);
        }

        [HttpPost]
        public ActionResult Branding(BrandingViewModel model)
        {
            AntiForgery.Validate();
            if (ModelState.IsValid)
            {
                // Get the original branding settings
                BrandingSettings original = PnPPartnerPackUtilities.GetTenantBrandingSettings();

                // If something changed in the branding settings compared with previous settings
                if (model.LogoImageUrl != original.LogoImageUrl ||
                    model.BackgroundImageUrl != original.BackgroundImageUrl ||
                    model.ColorFileUrl != original.ColorFileUrl ||
                    model.FontFileUrl != original.FontFileUrl ||
                    model.CSSOverrideUrl != original.CSSOverrideUrl ||
                    model.UICustomActionsUrl != original.UICustomActionsUrl)
                {
                    var newBrandingSettings = new BrandingSettings {
                        LogoImageUrl = model.LogoImageUrl,
                        BackgroundImageUrl = model.BackgroundImageUrl,
                        ColorFileUrl = model.ColorFileUrl,
                        FontFileUrl = model.FontFileUrl,
                        CSSOverrideUrl = model.CSSOverrideUrl,
                        UICustomActionsUrl = model.UICustomActionsUrl,
                        UpdatedOn = DateTime.Now,
                    };

                    // Update the last date and time of update
                    model.UpdatedOn = newBrandingSettings.UpdatedOn;

                    // Get the JSON representation of the settings
                    var jsonBranding = JsonConvert.SerializeObject(newBrandingSettings);

                    // Store the new tenant-wide settings in the Infrastructural Site Collection
                    PnPPartnerPackUtilities.SetPropertyBagValueToInfrastructure(
                        PnPPartnerPackConstants.PropertyBag_Branding, jsonBranding);
                }

                if (model.RollOut)
                {
                    // Create the asynchronous job
                    var job = new BrandingJob
                    {
                        Owner = ClaimsPrincipal.Current.Identity.Name,
                        Title = "Tenant Wide Branding",
                    };

                    // Enqueue the job for execution
                    model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);
                }
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult RefreshSites()
        {
            RefreshSitesViewModel model = new RefreshSitesViewModel();

            // Retrieve any pending, running, or failed job
            var runningJobs = ProvisioningRepositoryFactory.Current.GetProvisioningJobs(
                ProvisioningJobStatus.Pending | ProvisioningJobStatus.Running |
                ProvisioningJobStatus.Failed, typeof(RefreshSitesJob).FullName, 
                false);

            if (runningJobs != null && runningJobs.Length > 0)
            {
                // Get the most recent job instance
                var lastJob = runningJobs.OrderByDescending(j => j.ScheduledOn).First();

                // Configure the model accordingly
                model.Status = lastJob.Status == ProvisioningJobStatus.Pending | lastJob.Status == ProvisioningJobStatus.Running ? RefreshJobStatus.Running :
                    (lastJob.Status == ProvisioningJobStatus.Failed ? RefreshJobStatus.Failed : RefreshJobStatus.Idle);

                if (model.Status == RefreshJobStatus.Failed)
                {
                    model.ErrorMessage = lastJob.ErrorMessage;
                }
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult RefreshSites(RefreshSitesViewModel model)
        {
            AntiForgery.Validate();
            if (ModelState.IsValid)
            {
                // Create the asynchronous job
                var job = new RefreshSitesJob
                {
                    Owner = ClaimsPrincipal.Current.Identity.Name,
                    Title = "Tenant Wide Update Templates",
                };

                // Enqueue the job for execution
                model.JobId = ProvisioningRepositoryFactory.Current.EnqueueProvisioningJob(job);

                model.Status = RefreshJobStatus.Scheduled;
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