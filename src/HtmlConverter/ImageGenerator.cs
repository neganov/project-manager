using HiQPdf;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace HtmlConverter
{
    public static class ImageGenerator
    {
        public static void GetPngFromHtml(Stream stream, string html)
        {
            HtmlToImage converter = new HtmlToImage();
            converter.BrowserWidth = 600;
            converter.TrimToBrowserWidth = false;
            Image[] images = converter.ConvertHtmlToImage(html, null);
            if(images.Length > 0)
            {
                images[0].Save(stream, ImageFormat.Png);
            }
            else
            {
                throw new ApplicationException(
                    "HiQPdf conversion returned no images.");
            }
        }
    }
}
