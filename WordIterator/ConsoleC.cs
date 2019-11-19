using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordIterator
{
    class ConsoleC
    {
        public static void Write(ConsoleColor foreground, string text)
        {
            Console.ForegroundColor = foreground;
            Console.Write(text);
        }
        public static void Write(ConsoleColor foreground, ConsoleColor background, string text)
        {
            Console.ForegroundColor = foreground;
            Console.BackgroundColor = background;
            Console.Write(text);
            Console.ResetColor();
        }
        public static void WriteLine(ConsoleColor foreground, string text)
        {
            Console.ForegroundColor = foreground;
            Console.WriteLine(text);
        }
        public static void WriteLine(ConsoleColor foreground, ConsoleColor background, string text)
        {
            Console.ForegroundColor = foreground;
            Console.BackgroundColor = background;
            Console.WriteLine(text);
            Console.ResetColor();
        }
    }
}
