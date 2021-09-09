using System;
using System.CommandLine;
using System.CommandLine.Invocation;

namespace JpegToWord
{
    internal class Program
    {
        private static int Main(string[] args)
        {
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
            rootCommand.Handler = CommandHandler.Create<string[], string, string, string, string>(Invoker.Execute);

            return rootCommand.InvokeAsync(args).Result;
        }
    }
}