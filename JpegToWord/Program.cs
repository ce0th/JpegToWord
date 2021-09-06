using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Diagnostics;
using System.Drawing;

namespace JpegToWord
{
    class Program
    {
        static void Main(string[] args)
        {
            var doc = new Document();
            var section = doc.AddSection();
            var intro = section.AddParagraph();

            for (int i = 0; i < args.Length; i++)
            {
                var list = $"List of files:\nFile-{i + 1}: {args[0]}";
                intro.AppendText(list);
            }

            intro.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            intro.Format.AfterSpacing = 15;
            intro.Format.BeforeSpacing = 20;

            foreach (var arg in args)
            {
                var paragraph = section.AddParagraph();
                var image = paragraph.AppendPicture((byte[])(new ImageConverter()).ConvertTo(Image.FromFile(@$"{arg}"), typeof(byte[])));
                image.VerticalAlignment = ShapeVerticalAlignment.Center;
                image.HorizontalAlignment = ShapeHorizontalAlignment.Center;
                image.Width = 500;
                image.Height = 500;
            }

            Console.WriteLine("Enter the output filename: ");
            var filename = Console.ReadLine();
            Console.WriteLine("Enter the path: ");
            var path = Console.ReadLine();

            doc.SaveToFile($"{path}//{filename}.docx", FileFormat.Docx);

            var process = new Process();
            process.StartInfo.UseShellExecute = true;
            process.StartInfo.FileName = $"{path}//{filename}.docx";
            process.Start();
        }
    }
}
