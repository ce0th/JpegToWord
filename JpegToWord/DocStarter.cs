using System;
using System.Diagnostics;

namespace JpegToWord
{
    public class DocStarter
    {
        public void StartDocument(string path, string filename)
        {
            Process process = new Process {StartInfo = {UseShellExecute = true, FileName = $"{path}//{filename}.docx"}};
            process.Start();

            if (!process.Start())
            {
                Console.WriteLine("Couldn't find the process so it never ran");
            }
        }
    }
}