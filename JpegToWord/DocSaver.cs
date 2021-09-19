using Spire.Doc;
using System;
using System.IO;
using static System.Environment;

namespace JpegToWord
{
    public class DocSaver
    {
        public static void SaveDoc(Document doc, string path, string filename)
        {
            if (File.Exists(path + @"\" + filename + ".docx"))
            {
                Console.WriteLine($"Filename '{filename}' already exist in the directory, quitting");
                Exit(-1);
            }
            else
            {
                doc.SaveToFile($"{path}//{filename}.docx", FileFormat.Docx);
            }
        }
    }
}