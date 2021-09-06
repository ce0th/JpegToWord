using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using System.IO;
using System.Diagnostics;

namespace JpegToWord
{
    class Program
    {

        static void Main(string[] args)
        {

            var doc = new Document();
            var section = doc.AddSection();

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
