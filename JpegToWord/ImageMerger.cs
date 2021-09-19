using Newtonsoft.Json;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using static System.Environment;

namespace JpegToWord
{
    public class ImageMerger
    {
        public void MergeImagesIntoDoc(string[] args, Document doc, string headerJson = null, string spacing = null)
        {
            Section section = doc.AddSection();

            AddHeader(section, headerJson);
            AddImage(args, section, spacing);
            AddTable(args, section);
        }

        private static void AddTable(IReadOnlyList<string> args, Section section)
        {
            Paragraph tableTitle = section.AddParagraph();

            tableTitle.Format.AfterSpacing = 20;
            tableTitle.Format.BeforeSpacing = 10;

            int fileCount = args.Count;

            TextRange title = tableTitle.AppendText($"Totally {fileCount} files:");

            title.CharacterFormat.FontName = "Calibri";
            title.CharacterFormat.TextColor = Color.FromArgb(32, 32, 32);
            title.CharacterFormat.FontSize = 12;
            title.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            Table table = section.AddTable(true);

            string[] header = {"#", "File path"};
            table.ResetCells(args.Count + 1, header.Length);

            TableRow headerRow = table.Rows[0];
            headerRow.IsHeader = true;
            headerRow.Height = 23;
            headerRow.RowFormat.BackColor = Color.FromArgb(238, 246, 252);

            PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
            table.PreferredWidth = width;

            for (int i = 0; i < header.Length; i++)
            {
                Paragraph tableParagraph = headerRow.Cells[i].AddParagraph();

                headerRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                tableParagraph.Format.HorizontalAlignment = HorizontalAlignment.Left;

                TextRange textRange = tableParagraph.AppendText(header[i]);

                textRange.CharacterFormat.FontName = "Calibri";
                textRange.CharacterFormat.FontSize = 12;
                textRange.CharacterFormat.TextColor = Color.FromArgb(3, 116, 116);
                textRange.CharacterFormat.Bold = true;
            }

            for (int i = 0; i < args.Count; i++)
            {
                TableRow dataRow = table.Rows[i + 1];

                dataRow.Height = 20;

                table.Rows[i].Cells[0].SetCellWidth(7, CellWidthType.Percentage);
                table.Rows[i].Cells[1].SetCellWidth(93, CellWidthType.Percentage);

                Paragraph p1 = dataRow.Cells[0].AddParagraph();
                Paragraph p2 = dataRow.Cells[1].AddParagraph();

                p1.Format.HorizontalAlignment = HorizontalAlignment.Left;
                p2.Format.HorizontalAlignment = HorizontalAlignment.Justify;

                TextRange textRange1 = p1.AppendText($"{i + 1}");
                TextRange textRange2 = p2.AppendText(args[i]);

                textRange1.CharacterFormat.FontName = "Calibri";
                textRange2.CharacterFormat.FontName = "Calibri";

                textRange1.CharacterFormat.FontSize = 10;
                textRange2.CharacterFormat.FontSize = 10;

                textRange1.CharacterFormat.TextColor = Color.FromArgb(33, 33, 33);
                textRange2.CharacterFormat.TextColor = Color.FromArgb(33, 33, 33);
            }
        }

        private static void AddHeader(ISection section, string headerJson)
        {
            if (string.IsNullOrEmpty(headerJson))
            {
                return;
            }

            if (!File.Exists(headerJson))
            {
                Console.WriteLine("Unable to find json, check the path, quitting");
                Exit(-1);
            }

            Paragraph intro = section.AddParagraph();

            intro.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            intro.Format.AfterSpacing = 10;
            intro.Format.BeforeSpacing = 10;
            intro.Format.LineSpacing = 9;

            string jsonString = File.ReadAllText(headerJson);

            Dictionary<string, string> dictionary =
                JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

            if (dictionary == null)
            {
                return;
            }

            for (int i = 0; i < dictionary.Count; i++)
            {
                (string key, string value) = dictionary.ElementAt(i);

                TextRange text = intro.AppendText(key + ": " + value + "\n");
                text.CharacterFormat.FontName = "Calibri";
                text.CharacterFormat.FontSize = 14;
                text.CharacterFormat.TextColor = Color.FromArgb(32, 32, 32);
            }
        }

        private static void AddImage(IEnumerable<string> args, ISection section, string spacing = null)
        {
            foreach (string arg in args)
            {
                Paragraph paragraph = section.AddParagraph();

                if (File.Exists(arg))
                {
                    DocPicture image = paragraph.AppendPicture(
                        (byte[])new ImageConverter().ConvertTo(Image.FromFile(@$"{arg}"), typeof(byte[])));
                    image.VerticalAlignment = ShapeVerticalAlignment.Center;
                    image.HorizontalAlignment = ShapeHorizontalAlignment.Center;

                    Image img = Image.FromFile(arg);

                    image.Width = 500;
                    image.Height = 500 * (img.Height / (float)img.Width);

                    paragraph.Format.BeforeSpacing = StringParser.ParseStringToInt(spacing);
                }
                else
                {
                    Console.WriteLine($"Check the file path: {arg}, quitting");

                    Exit(-1);
                }
            }
        }
    }
}