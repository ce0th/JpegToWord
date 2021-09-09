using System;
using Spire.Doc;

namespace JpegToWord
{
    internal static class MergeUtil
    {
        public static void MergeImagesIntoDoc(string[] images, string output, string filename)
        {
            Console.WriteLine("Building doc with no header ...");

            var doc = new Document();
            var dc = new DocCreator();
            dc.CreateWordDoc(images, doc);

            var saver = new DocSaver();
            saver.SaveDoc(doc, output, filename);

            var starter = new DocStarter();
            starter.StartDocument(output, filename);
        }

        public static void MergeImagesIntoDocWithHeader(string[] images, string output, string filename, string header)
        {
            Console.WriteLine("Building doc with header ...");

            var doc = new Document();
            var dc = new DocCreator();
            dc.CreateWordDocWithHeader(images, doc, header);

            var saver = new DocSaver();
            saver.SaveDoc(doc, output, filename);

            var starter = new DocStarter();
            starter.StartDocument(output, filename);
        }
    }
}