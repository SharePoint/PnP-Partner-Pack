using OfficeDevPnP.PartnerPack.SiteProvisioning.Components;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Binders
{
    public class PrincipalViewModelBinder : DefaultModelBinder
    {
        public override object BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
        {
            var model = (PrincipalViewModel)base.BindModel(controllerContext, bindingContext);

            if (!String.IsNullOrWhiteSpace(model.Mail))
            {
                var result = UserUtility.GetUser("users/" + model.Mail);
                model.DisplayName = result?.DisplayName;
            }

            return model;
        }
    }
}