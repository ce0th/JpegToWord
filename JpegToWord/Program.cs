using Spire.Doc;

namespace JpegToWord
{
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