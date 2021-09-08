using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using Spire.Doc;

namespace JpegToWord
{
    internal class Program
    {
        public static readonly List<string> ImageExtensions = new List<string> {".JPG", ".JPE", ".BMP", ".GIF", ".PNG"};

        private static int Main(string[] args)
        {
            var rootCommand = new RootCommand
            {
                new Option<string[]>(
                    "--images",
                    "Specify paths to your incoming images"),
                new Option<string>(
                    "--imageFolder",
                    "Specify path to your folder containing incoming images"),
                new Option<string>(
                    "--filename",
                    description: "Name for output Word file\n",
                    getDefaultValue: () => $"MergedFile{DateTime.Now.ToString("yyMMddHHmmssff")}"
                ),
                new Option<string>(
                    "--output",
                    description: "Path to directory where the output Word will be created\n",
                    getDefaultValue: () => Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                ),
                new Option<string>(
                    "--header",
                    "Specify path to your Json file")
            };

            rootCommand.Description = "Console App to merge image files into one Word document";
            rootCommand.Handler = CommandHandler.Create<string[], string, string, string, string>(Execute);

            return rootCommand.InvokeAsync(args).Result;
        }

        public static void Execute(string[] images, string imageFolder, string output, string filename, string header)
        {
            if (images.Length == 0)
                Console.WriteLine("No images in args, checking imageFolder ...");
            else
                MergeImagesIntoDoc(images, output, filename);

            if (string.IsNullOrEmpty(imageFolder))
            {
                Console.WriteLine("imageFolder not provided ...");
            }
            else
            {
                var filePaths = CheckImages(imageFolder);
                MergeImagesIntoDoc(filePaths, output, filename);
            }

            if (string.IsNullOrEmpty(output))
                Console.WriteLine($"No output specified, saving to {output ?? "null"}");
            else if ((images.Length != 0 || !string.IsNullOrEmpty(imageFolder)) &&
                     !string.IsNullOrEmpty(output)) Console.WriteLine($"Output directory is  {output ?? "null"}");

            if (string.IsNullOrEmpty(header))
            {
                Console.WriteLine("No Json path is provided ...");
            }
            else
            {
                if (string.IsNullOrEmpty(imageFolder) && images.Length != 0)
                    MergeImagesIntoDoc(images, output, filename, header);

                if (!string.IsNullOrEmpty(imageFolder) && images.Length == 0)
                {
                    var filePaths = CheckImages(imageFolder);
                    MergeImagesIntoDoc(filePaths, output, filename, header);
                }
            }

            if (images.Length == 0 && string.IsNullOrEmpty(imageFolder))
                Console.WriteLine("No images were provided! Exiting ...");
        }

        public static void MergeImagesIntoDoc(string[] images, string output, string filename)
        {
            var doc = new Document();
            var dc = new DocCreator();
            dc.CreateWordDoc(images, doc);

            var saver = new DocSaver();
            saver.SaveDoc(doc, output, filename);

            var starter = new DocStarter();
            starter.StartDocument(output, filename);
        }

        public static void MergeImagesIntoDoc(string[] images, string output, string filename, string header)
        {
            var doc = new Document();
            var dc = new DocCreator();
            dc.CreateWordDoc(images, doc, header);

            var saver = new DocSaver();
            saver.SaveDoc(doc, output, filename);

            var starter = new DocStarter();
            starter.StartDocument(output, filename);
        }

        public static string[] CheckImages(string imageFolder)
        {
            var filePaths = Directory.GetFiles(imageFolder);

            return filePaths.Where(file => ImageExtensions.Contains(Path.GetExtension(file).ToUpperInvariant()))
                .ToArray();
        }
    }
}