using System;
using System.CommandLine;
using System.CommandLine.Invocation;

namespace JpegToWord
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine(" Show help and usage information: -?, -h, --help  ");
                return -1;
            }

            RootCommand rootCommand = new RootCommand
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
                    getDefaultValue: () => $"MergedFile{DateTime.Now:yyMMddHHmmssff}"
                ),
                new Option<string>(
                    "--output",
                    description: "Path to directory where the output Word will be created\n",
                    getDefaultValue: () => Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                ),
                new Option<string>(
                    "--run",
                    description: "Specify true if want to run file after creation, default is false",
                    getDefaultValue: () => null
                ),
                new Option<string>(
                    "--spacing",
                    description: "Specify spacing between images, default is 0",
                    getDefaultValue: () => null
                ),
                new Option<string>(
                    "--header",
                    "Specify path to your Json file")
            };

            rootCommand.Description = "Console App to merge image files into one Word document";
            rootCommand.Handler =
                CommandHandler.Create<string[], string, string, string, string, string, string>(Invoker.Execute);

            return rootCommand.InvokeAsync(args).Result;
        }
    }
}