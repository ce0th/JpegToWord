using System;

public class FilePathAndName
{
    public string GetFileName()
    {
        Console.WriteLine("Enter the output filename: ");
        return Console.ReadLine();
    }

    public string GetFilePath()
    {
        Console.WriteLine("Enter the path: ");
        return Console.ReadLine();
    }
}