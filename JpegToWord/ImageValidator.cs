using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JpegToWord
{
    internal static class ImageValidator
    {
        private static readonly List<string> ImageExtensions = new List<string>
        {
            ".JPG"
            //  ".JPE",
            //  ".BMP",
            //  ".GIF",
            //  ".PNG"
        };

        public static string[] IsImage(string imageFolder = null, string[] images = null)
        {
            if (!string.IsNullOrEmpty(imageFolder))
            {
                string[] filePaths = Directory.GetFiles(imageFolder);

                if (filePaths.Length != filePaths.Where(file =>
                        ImageExtensions.Contains(Path.GetExtension(file).ToUpperInvariant()))
                    .ToArray().Length)
                {
                    Console.WriteLine("Some files in directory was not of image type so excluded from merge");
                }

                return filePaths.Where(file => ImageExtensions.Contains(Path.GetExtension(file).ToUpperInvariant()))
                    .ToArray();
            }

            if (images != null)
            {
                if (images.Length != images
                    .Where(file => ImageExtensions.Contains(Path.GetExtension(file).ToUpperInvariant()))
                    .ToArray().Length)
                {
                    Console.WriteLine(
                        "Some files from images arguments was excluded from merge because are not of image type");
                }

                return images.Where(file => ImageExtensions.Contains(Path.GetExtension(file).ToUpperInvariant()))
                    .ToArray();
            }


            return new string[] { };
        }
    }
}