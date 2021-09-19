using System;
using System.IO;

namespace JpegToWord
{
    internal static class Invoker
    {
        public static void Execute(
            string[] images,
            string imageFolder,
            string output,
            string filename,
            string header = null,
            string run = null,
            string spacing = null
        )
        {
            DocCreator docCreator = new DocCreator();

            if (images.Length == 0 && string.IsNullOrEmpty(imageFolder))
            {
                Console.WriteLine("No images were provided, quitting ...");
                return;
            }

            if (string.IsNullOrEmpty(imageFolder))
            {
                if (images.Length > 350)
                {
                    Console.WriteLine("Can't process more than 350 images, quitting ...");
                    return;
                }

                string[] filePaths = ImageValidator.IsImage(null, images);
                docCreator.MergeImagesIntoDoc(filePaths, output, filename, header, spacing, run);
            }
            else
            {
                if (Directory.GetFiles(imageFolder, "*", SearchOption.TopDirectoryOnly).Length > 350)
                {
                    Console.WriteLine("Can't process more than 350 images, quitting ...");
                    return;
                }

                string[] filePaths = ImageValidator.IsImage(imageFolder);
                docCreator.MergeImagesIntoDoc(filePaths, output, filename, header, spacing, run);
            }
        }
    }
}