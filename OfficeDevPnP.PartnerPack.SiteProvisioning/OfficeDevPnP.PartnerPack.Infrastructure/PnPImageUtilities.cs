using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure
{
    public static class PnPImageUtilities
    {
        public static Stream FixedSizeImageStream(this Stream imageStream, Int32 width, Int32 height)
        {
            Image sourceImage = Image.FromStream(imageStream);
            MemoryStream mem = new MemoryStream();
            FixedSize(sourceImage, width, height).Save(mem, ImageFormat.Png);
            mem.Position = 0;
            return (mem);
        }

        private static Image FixedSize(Image imgPhoto, Int32 width, Int32 height)
        {
            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;
            int sourceX = 0;
            int sourceY = 0;
            int destX = 0;
            int destY = 0;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)width / (float)sourceWidth);
            nPercentH = ((float)height / (float)sourceHeight);
            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
                destX = System.Convert.ToInt16((width -
                              (sourceWidth * nPercent)) / 2);
            }
            else
            {
                nPercent = nPercentW;
                destY = System.Convert.ToInt16((height -
                              (sourceHeight * nPercent)) / 2);
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap bmPhoto = new Bitmap(width, height,
                              PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                             imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.Gray);
            grPhoto.InterpolationMode =
                    InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new Rectangle(destX, destY, destWidth, destHeight),
                new Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            return bmPhoto;
        }
    }
}
