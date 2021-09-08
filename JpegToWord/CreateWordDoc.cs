using System.Drawing;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

public class DocCreator
{
    public void CreateWordDoc(string[] args, Document doc)
    {
        var section = doc.AddSection();


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

    public void CreateWordDoc(string[] args, Document doc, string header)
    {
        var section = doc.AddSection();

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

        var summary = section.AddParagraph();
        summary.Format.HorizontalAlignment = HorizontalAlignment.Justify;
        summary.Format.AfterSpacing = 10;
        summary.Format.BeforeSpacing = 10;

        var jsonString = File.ReadAllText(header);
        var jsonUtil = new JsonUtil();
        var prettified = jsonUtil.JsonPrettify(jsonString);
        var text = summary.AppendText(prettified);
        text.CharacterFormat.FontName = "Cambria";
        text.CharacterFormat.FontSize = 14;
        text.CharacterFormat.TextColor = Color.FromArgb(37, 40, 95);
    }
}