using Spire.Doc;
using System;

namespace JpegToWord
{
    internal static class MergeUtil
    {
        public static void MergeImagesIntoDoc(string[] images, string output, string filename)
        {
            Console.WriteLine("Building doc with no header ...");

            Document doc = new Document();
            DocCreator dc = new DocCreator();
            dc.CreateWordDoc(images, doc);

            DocSaver saver = new DocSaver();
            saver.SaveDoc(doc, output, filename);

            DocStarter starter = new DocStarter();
            starter.StartDocument(output, filename);
        }

        public static void MergeImagesIntoDocWithHeader(string[] images, string output, string filename, string header)
        {
            Console.WriteLine("Building doc with header ...");

            Document doc = new Document();
            DocCreator dc = new DocCreator();
            dc.CreateWordDocWithHeader(images, doc, header);

            DocSaver saver = new DocSaver();
            saver.SaveDoc(doc, output, filename);

            DocStarter starter = new DocStarter();
            starter.StartDocument(output, filename);
        }
    }
}