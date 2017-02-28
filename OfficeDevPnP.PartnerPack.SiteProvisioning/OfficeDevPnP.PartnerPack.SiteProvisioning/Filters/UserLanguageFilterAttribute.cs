using OfficeDevPnP.PartnerPack.SiteProvisioning.Components;
using System;
using System.Globalization;
using System.Threading;
using System.Web;
using System.Web.Mvc;

public class UserLanguageFilterAttribute : ActionFilterAttribute
{

    public override void OnActionExecuting(ActionExecutingContext filterContext)
    {
        var userLanguage = GetUserPreferredLanguage(filterContext.HttpContext);

        var culture = new CultureInfo(userLanguage);
        Thread.CurrentThread.CurrentCulture = culture;
        Thread.CurrentThread.CurrentUICulture = culture;
    }

    public static string GetUserPreferredLanguage(HttpContextBase httpContext)
    {
        //Default language
        var result = "en-us";

        var cookie = default(HttpCookie);

        if (httpContext == null)
            return result;

        try
        {
            cookie = httpContext.Request?.Cookies.Get("PnPUserInfo");
            if (cookie == default(HttpCookie) || string.IsNullOrEmpty(cookie["language"]))
            {
                var user = UserUtility.GetCurrentUser();
                cookie = new HttpCookie("PnPUserInfo");
                cookie["language"] = user.PreferredLanguage;
                cookie.Expires = DateTime.Now.AddMonths(1);

                httpContext.Response?.Cookies.Add(cookie);
            }

            result = cookie["language"];

        }
        catch (Exception)
        {

        }


        return result;
    }
}