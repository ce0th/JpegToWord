namespace JpegToWord
{
    internal static class StringParser
    {
        public static int ParseStringToInt(string spacing)
        {
            int numValue = int.Parse(spacing);

            if (numValue < 0 || numValue > 100)
            {
                return 0;
            }

            return numValue;
        }
    }
}