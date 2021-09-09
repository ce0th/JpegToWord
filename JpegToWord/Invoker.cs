using System;

namespace JpegToWord
{
    internal static class Invoker
    {
        public static void Execute(string[] images, string imageFolder, string output, string filename, string header)
        {
            if (string.IsNullOrEmpty(header))
            {
                Console.WriteLine("No Json path is provided ...");

                if (images.Length == 0)
                {
                    Console.WriteLine("No images in arguments, checking imageFolder ...");

                    if (string.IsNullOrEmpty(imageFolder))
                    {
                        Console.WriteLine("imageFolder not provided ...");
                    }
                    else
                    {
                        var filePaths = ImageUtil.CheckImages(imageFolder);
                        MergeUtil.MergeImagesIntoDoc(filePaths, output, filename);
                    }
                }

                else
                {
                    MergeUtil.MergeImagesIntoDoc(images, output, filename);
                }
            }
            else
            {
                if (string.IsNullOrEmpty(imageFolder) && images.Length > -0)
                    MergeUtil.MergeImagesIntoDocWithHeader(images, output, filename, header);

                if (!string.IsNullOrEmpty(imageFolder) && images.Length == 0)
                {
                    var filePaths = ImageUtil.CheckImages(imageFolder);
                    MergeUtil.MergeImagesIntoDocWithHeader(filePaths, output, filename, header);
                }
            }

            if (string.IsNullOrEmpty(output))
                Console.WriteLine($"No output specified, saving to {output ?? "null"}");
            else if ((images.Length != 0 || !string.IsNullOrEmpty(imageFolder)) &&
                     !string.IsNullOrEmpty(output)) Console.WriteLine($"Output directory is  {output ?? "null"}");

            if (images.Length == 0 && string.IsNullOrEmpty(imageFolder))
                Console.WriteLine("No images were provided! Exiting ...");
        }
    }
}