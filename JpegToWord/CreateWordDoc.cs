using Newtonsoft.Json;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

public class DocCreator
{
    public void CreateWordDoc(string[] args, Document doc)
    {
        var section = doc.AddSection();


        foreach (var arg in args)
        {
            var paragraph = section.AddParagraph();
            var image = paragraph.AppendPicture(
                (byte[])new ImageConverter().ConvertTo(Image.FromFile(@$"{arg}"), typeof(byte[])));
            image.VerticalAlignment = ShapeVerticalAlignment.Center;
            image.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            image.Width = 500;
            image.Height = 500;
        }

        var footer = section.AddParagraph();

        for (var i = 0; i < args.Length; i++)
        {
            var list = $"File-{i + 1}: {args[0]}";
            footer.AppendText(list);
            footer.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            footer.Format.AfterSpacing = 10;
            footer.Format.BeforeSpacing = 10;
        }
    }

    public void CreateWordDocWithHeader(string[] args, Document doc, string headerJson)
    {
        var section = doc.AddSection();

        var header = section.AddParagraph();
        header.Format.HorizontalAlignment = HorizontalAlignment.Justify;
        header.Format.AfterSpacing = 10;
        header.Format.BeforeSpacing = 10;

        var jsonString = File.ReadAllText(headerJson);

        Dictionary<string, string> dictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

        for (int count = 0; count < dictionary.Count; count++)
        {
            var element = dictionary.ElementAt(count);
            var Key = element.Key;
            var Value = element.Value;
            var text = header.AppendText(Key + ": " + Value + "\n");
            text.CharacterFormat.FontName = "Cambria";
            text.CharacterFormat.FontSize = 14;
            text.CharacterFormat.TextColor = Color.FromArgb(37, 40, 95);
        }

        foreach (var arg in args)
        {
            var paragraph = section.AddParagraph();
            var image = paragraph.AppendPicture(
                (byte[])new ImageConverter().ConvertTo(Image.FromFile(@$"{arg}"), typeof(byte[])));
            image.VerticalAlignment = ShapeVerticalAlignment.Center;
            image.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            image.Width = 500;
            image.Height = 500;
        }

        var footer = section.AddParagraph();

        for (var i = 0; i < args.Length; i++)
        {
            var list = $"File-{i + 1}: {args[0]}\n";
            footer.AppendText(list);
            footer.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            footer.Format.AfterSpacing = 10;
            footer.Format.BeforeSpacing = 10;

        }
    }
}