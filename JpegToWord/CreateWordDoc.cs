using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;

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