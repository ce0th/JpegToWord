using Spire.Doc;

namespace JpegToWord
{
    public class DocSaver
    {
        public void SaveDoc(Document doc, string path, string filename)
        {
            doc.SaveToFile($"{path}//{filename}.docx", FileFormat.Docx);
        }
    }
}