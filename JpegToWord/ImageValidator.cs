using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JpegToWord
{
    internal static class ImageValidator
    {
        public static readonly List<string> ImageExtensions = new List<string>
        {
            ".JPG"
            //  ".JPE",
            //  ".BMP",
            //  ".GIF",
            //  ".PNG"
        };

        public static string[] CheckImages(string imageFolder)
        {
            string[] filePaths = Directory.GetFiles(imageFolder);

            return filePaths.Where(file => ImageExtensions.Contains(Path.GetExtension(file).ToUpperInvariant()))
                .ToArray();
        }
    }
}