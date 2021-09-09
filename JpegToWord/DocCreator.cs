using Newtonsoft.Json;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace JpegToWord
{
    public class DocCreator
    {
        public void CreateWordDoc(string[] args, Document doc)
        {
            Section section = doc.AddSection();

            foreach (string arg in args)
            {
                Paragraph paragraph = section.AddParagraph();
                DocPicture image = paragraph.AppendPicture(
                    (byte[])new ImageConverter().ConvertTo(Image.FromFile(@$"{arg}"), typeof(byte[])));
                image.VerticalAlignment = ShapeVerticalAlignment.Center;
                image.HorizontalAlignment = ShapeHorizontalAlignment.Center;
                Image img = Image.FromFile(arg);
                image.Width = 500;
                image.Height = 500 * (img.Height / img.Width);
            }

            Paragraph footer = section.AddParagraph();

            for (int i = 0; i < args.Length; i++)
            {
                string list = $"File-{i + 1}: {args[0]}";
                footer.AppendText(list);
                footer.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                footer.Format.AfterSpacing = 10;
                footer.Format.BeforeSpacing = 10;
            }
        }

        public void CreateWordDocWithHeader(string[] args, Document doc, string headerJson)
        {
            Section section = doc.AddSection();

            Paragraph header = section.AddParagraph();
            header.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            header.Format.AfterSpacing = 10;
            header.Format.BeforeSpacing = 10;

            string jsonString = File.ReadAllText(headerJson);

            Dictionary<string, string>? dictionary =
                JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

            for (int count = 0; count < dictionary.Count; count++)
            {
                KeyValuePair<string, string> element = dictionary.ElementAt(count);
                string Key = element.Key;
                string Value = element.Value;
                TextRange text = header.AppendText(Key + ": " + Value + "\n");
                text.CharacterFormat.FontName = "Cambria";
                text.CharacterFormat.FontSize = 14;
                text.CharacterFormat.TextColor = Color.FromArgb(37, 40, 95);
            }

            foreach (string arg in args)
            {
                Paragraph paragraph = section.AddParagraph();
                DocPicture image = paragraph.AppendPicture(
                    (byte[])new ImageConverter().ConvertTo(Image.FromFile(arg), typeof(byte[])));
                image.VerticalAlignment = ShapeVerticalAlignment.Center;
                image.HorizontalAlignment = ShapeHorizontalAlignment.Center;
                Image img = Image.FromFile(arg);
                image.Width = 500;
                image.Height = 500 * (img.Height / img.Width);
            }


            Paragraph footer = section.AddParagraph();

            for (int i = 0; i < args.Length; i++)
            {
                string list = $"File-{i + 1}: {args[0]}\n";
                footer.AppendText(list);
                footer.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                footer.Format.AfterSpacing = 10;
                footer.Format.BeforeSpacing = 10;
            }
        }
    }
}