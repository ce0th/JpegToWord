using System;

namespace JpegToWord
{
    internal static class Invoker
    {
        public static void Execute(string[] images, string imageFolder, string output, string filename,
            string header = null,
            string run = null, string spacing = null)
        {
            DocCreator docCreator = new DocCreator();

            if (images.Length == 0 && string.IsNullOrEmpty(imageFolder))
            {
                Console.WriteLine("No images were provided, quitting ...");
                return;
            }

            if (string.IsNullOrEmpty(imageFolder))
            {
                docCreator.MergeImagesIntoDoc(images, output, filename, header, spacing, run);
            }
            else
            {
                string[] filePaths = ImageValidator.IsImage(imageFolder);
                docCreator.MergeImagesIntoDoc(filePaths, output, filename, header, spacing, run);
            }
        }
    }
}