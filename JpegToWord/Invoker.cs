using System;

namespace JpegToWord
{
    internal static class Invoker
    {
        public static void Execute(string[] images, string imageFolder, string output, string filename, string header)
        {
            if (string.IsNullOrEmpty(output))
            {
                Console.WriteLine($"No output specified, saving to {output ?? "null"}");
            }

            if (!string.IsNullOrEmpty(output))
            {
                Console.WriteLine($"Output directory is  {output ?? "null"}");
            }

            if (string.IsNullOrEmpty(header))
            {
                Console.WriteLine("No Json path is provided ...");

                if (images.Length == 0 || !string.IsNullOrEmpty(imageFolder))
                {
                    Console.WriteLine("No images in arguments, loading from imageFolder ...");

                    string[] filePaths = ImageUtil.CheckImages(imageFolder);
                    MergeUtil.MergeImagesIntoDoc(filePaths, output, filename);
                }

                if (images.Length > 0 && string.IsNullOrEmpty(imageFolder))
                {
                    MergeUtil.MergeImagesIntoDoc(images, output, filename);
                }
                else
                {
                    Console.WriteLine("No images were provided! Exiting ...");
                }
            }
            else
            {
                Console.WriteLine("Json path is provided ...");

                if (images.Length == 0 || !string.IsNullOrEmpty(imageFolder))
                {
                    Console.WriteLine("Loading from imageFolder ... ");

                    string[] filePaths = ImageUtil.CheckImages(imageFolder);
                    MergeUtil.MergeImagesIntoDocWithHeader(filePaths, output, filename, header);
                }

                if (images.Length > 0 || string.IsNullOrEmpty(imageFolder))
                {
                    Console.WriteLine("Loading from images path ..." +
                                      "");
                    MergeUtil.MergeImagesIntoDocWithHeader(images, output, filename, header);
                }

                if (images.Length == 0 && string.IsNullOrEmpty(imageFolder))

                {
                    Console.WriteLine("No images were provided! Exiting ...");
                }
            }
        }
    }
}