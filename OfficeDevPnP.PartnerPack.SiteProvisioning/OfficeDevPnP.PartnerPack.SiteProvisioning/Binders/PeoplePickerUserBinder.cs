using OfficeDevPnP.PartnerPack.SiteProvisioning.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Binders
{
    public class PeoplePickerUserBinder : IModelBinder
    {
        public object BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
        {
            var v = bindingContext.ValueProvider.GetValue(bindingContext.ModelName);
            if (v != null && v.RawValue != null)
            {
                return Newtonsoft.Json.JsonConvert.DeserializeObject<PeoplePickerUser[]>(v.AttemptedValue);
            }

            return null;
        }
    }

}