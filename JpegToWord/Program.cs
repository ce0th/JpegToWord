using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;

namespace JpegToWord
{
    public class DocCreator
    {
        public void CreateWordDoc(string[] args, Document doc)
        {
            var section = doc.AddSection();
            var intro = section.AddParagraph();

            for (var i = 0; i < args.Length; i++)
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
                var image = paragraph.AppendPicture(
                    (byte[]) new ImageConverter().ConvertTo(Image.FromFile(@$"{arg}"), typeof(byte[])));
                image.VerticalAlignment = ShapeVerticalAlignment.Center;
                image.HorizontalAlignment = ShapeHorizontalAlignment.Center;
                image.Width = 500;
                image.Height = 500;
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            var doc = new Document();

            var dc = new DocCreator();
            dc.CreateWordDoc(args, doc);

            var file = new FilePathAndName();
            var filename = file.GetFileName();
            var path = file.GetFilePath();

            var saver = new DocSaver();
            saver.SaveDoc(doc, path, filename);

            var starter = new DocStarter();
            starter.StartDocument(path, filename);
        }
    }
}