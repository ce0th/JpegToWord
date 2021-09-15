using System;

namespace JpegToWord
{
    internal static class Invoker
    {
        public static void Execute(string[] images, string imageFolder, string output, string filename, string header,
            string run, string spacing)
        {
            if (images.Length == 0 && string.IsNullOrEmpty(imageFolder))
            {
                Console.WriteLine("No images were provided! Exiting ...");
                return;
            }

            if (string.IsNullOrEmpty(header))
            {
                Console.WriteLine("No Json path is provided ...");

                if (images.Length == 0 && !string.IsNullOrEmpty(imageFolder))
                {
                    Console.WriteLine("No image paths in arguments, merging from imageFolder ...");

                    string[] filePaths = ImageValidator.CheckImages(imageFolder);

                    if (string.IsNullOrEmpty(run))
                    {
                        DocCreator.MergeImagesIntoDoc(filePaths, output, filename, null, spacing);
                    }

                    else
                    {
                        DocCreator.MergeImagesIntoDoc(filePaths, output, filename, null, spacing, run);
                    }
                }

                if (images.Length > 0 && string.IsNullOrEmpty(imageFolder))
                {
                    if (string.IsNullOrEmpty(run))
                    {
                        DocCreator.MergeImagesIntoDoc(images, output, filename, null, spacing);
                    }
                    else
                    {
                        DocCreator.MergeImagesIntoDoc(images, output, filename, null, spacing, run);
                    }
                }
            }
            else
            {
                Console.WriteLine("Json path is provided ...");

                if (images.Length == 0 && !string.IsNullOrEmpty(imageFolder))
                {
                    Console.WriteLine("Loading from imageFolder ... ");

                    string[] filePaths = ImageValidator.CheckImages(imageFolder);

                    if (string.IsNullOrEmpty(run))
                    {
                        DocCreator.MergeImagesIntoDoc(filePaths, output, filename, header);
                    }
                    else
                    {
                        DocCreator.MergeImagesIntoDoc(filePaths, output, filename, header, null, run);
                        Console.WriteLine("Run");
                    }
                }

                if (images.Length > 0 && string.IsNullOrEmpty(imageFolder))
                {
                    Console.WriteLine("Loading from images path ...");

                    if (string.IsNullOrEmpty(run))
                    {
                        DocCreator.MergeImagesIntoDoc(images, output, filename, header);
                    }
                    else
                    {
                        DocCreator.MergeImagesIntoDoc(images, output, filename, header, null, run);
                    }
                }
            }
        }
    }
}