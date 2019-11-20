using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordIterator
{
    class StringMethods
    {
        // Truncates a string to the given length, or returns the whole string if too short.
        public static string Shorten(string inputText, int maxLength)
        {
            int l = inputText.Length;
            int length = (l <= maxLength ? l : maxLength);
            string newString = inputText.Substring(0, length);
            return newString;
        }

        // The same as Shorten(), but removes line-breaks and excess white-space first.
        public static string TrimAndShorten(string inputText, int maxLength)
        {
            return Shorten(inputText.Replace("/n", " ").Replace("/r", " ").Trim(), maxLength);
        }
    }
}
