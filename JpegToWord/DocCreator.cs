using Spire.Doc;
using System;

namespace JpegToWord
{
    internal static class DocCreator
    {
        public static void MergeImagesIntoDoc(string[] images, string output, string filename, string header = null,
            string spacing = null, string run = null)
        {
            Document doc = new Document();
            ImageMerger dc = new ImageMerger();
            dc.MergeImagesIntoDoc(images, doc, header, spacing);

            DocSaver saver = new DocSaver();
            saver.SaveDoc(doc, output, filename);

            Console.WriteLine("Document was saved ...");

            if (string.IsNullOrEmpty(run))
            {
                return;
            }

            DocStarter starter = new DocStarter();
            starter.StartDocument(output, filename);
        }
    }
}