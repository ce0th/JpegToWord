using System;
using static System.Environment;

namespace JpegToWord
{
    internal static class StringParser
    {
        public static int ParseStringToInt(string spacing)
        {
            if (string.IsNullOrEmpty(spacing))
            {
                return 0;
            }

            try
            {
                int numValue = int.Parse(spacing);

                if (numValue > 0 && numValue < 100)
                {
                    return numValue;
                }
            }
            catch (FormatException)
            {
                Console.WriteLine($"Unable to parse spacing {spacing}, quitting ");
                Exit(-1);
            }

            return 0;
        }
    }
}