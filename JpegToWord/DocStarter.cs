using System.Diagnostics;

public class DocStarter
{
    public void StartDocument(string path, string filename)
    {
        var process = new Process {StartInfo = {UseShellExecute = true, FileName = $"{path}//{filename}.docx"}};
        process.Start();
    }
}