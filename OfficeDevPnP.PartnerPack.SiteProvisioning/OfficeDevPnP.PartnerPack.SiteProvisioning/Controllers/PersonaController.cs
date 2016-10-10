using Newtonsoft.Json;
using OfficeDevPnP.PartnerPack.Infrastructure;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Components;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Controllers
{
    [Authorize]
    public class PersonaController : Controller
    {
        public ActionResult GetPhoto(String upn, Int32 width = 0, Int32 height = 0)
        {
            Stream result = null;
            String contentType = "image/png";

            var sourceStream = GetUserPhoto(upn);

            if (sourceStream != null && width != 0 && height != 0)
            {
                Image sourceImage = Image.FromStream(sourceStream);
                Image resultImage = ScaleImage(sourceImage, width, height);

                result = new MemoryStream();
                resultImage.Save(result, ImageFormat.Png);
                result.Position = 0;
            }
            else
            {
                result = sourceStream;
            }

            if (result != null)
            {
                return base.File(result, contentType);
            }
            else
            {
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.NoContent);
            }
        }

        [HttpGet]
        public async Task<ActionResult> SearchPeopleOrGroups(String searchText, Boolean searchGroups)
        {
            PeoplePickerHelper helper = new PeoplePickerHelper();
            var results = await helper.GetPeoplePickerSearchData(searchText, searchGroups);

            return Json(results, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// This method retrieves the photo of a single user from Azure AD
        /// </summary>
        /// <param name="upn">The UPN of the user</param>
        /// <returns>The user's photo retrieved from Azure AD</returns>
        private static Stream GetUserPhoto(String upn)
        {
            String contentType = "image/png";
            Stream result = null;

            try
            {
                result = HttpHelper.MakeGetRequestForStream(
                    String.Format("{0}users/{1}/photo/$value",
                        MicrosoftGraphConstants.MicrosoftGraphV1BaseUri, upn),
                    contentType,
                    MicrosoftGraphHelper.GetAccessTokenForCurrentUser(MicrosoftGraphConstants.MicrosoftGraphResourceId));
            }
            catch (Exception)
            {
                // Ignore any exception related to image download, 
                // in order to get the default icon with user's initials
            }

            return (result);
        }

        private Image ScaleImage(Image image, int maxWidth, int maxHeight)
        {
            var ratioX = (double)maxWidth / image.Width;
            var ratioY = (double)maxHeight / image.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(image.Width * ratio);
            var newHeight = (int)(image.Height * ratio);

            var newImage = new Bitmap(newWidth, newHeight);

            using (var graphics = Graphics.FromImage(newImage))
                graphics.DrawImage(image, 0, 0, newWidth, newHeight);

            return newImage;
        }
    }
}