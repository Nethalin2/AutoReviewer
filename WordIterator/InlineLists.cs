using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace WordIterator
{
    class InlineLists
    {
        public static bool IsMatch(String textToCheck)
        {
            string[] patterns = {
                @".*\bi\..*\bii\..*", // Matches i. ii.
                @".*\ba\..*\bb\..*", // Matches a. b.
                @".*\b1\..*\b2\..*", // Matches 1. 2.
                @".*\bi\).*\bii\).*", // Matches i) ii)
                @".*\ba\).*\bb\).*", // Matches a) b)
                @".*\b1\).*\b2\).*", // Matches 1) 2)
                @".*\(i\).*\(ii\).*", // Matches (i) (ii)
                @".*\(a\).*\(b\).*", // Matches (a) (b)
                @".*\(1\).*\(2\).*" // Matches (1) (2)
            };
            foreach(string pattern in patterns)
            {
                if (Regex.IsMatch(textToCheck, pattern, RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }


        public static void DetectAll(Document doc)
        {
            ConsoleC.WriteLine(ConsoleColor.White, "Checking every paragraph for inline lists...");
            foreach (Paragraph para in doc.Paragraphs)
            {
                if (IsMatch(para.Range.Text))
                {
                    Comments.Add(doc, para, "This paragraph seems to contain a list. Consider rephrasing as a bulleted or numbered list.");
                    ConsoleC.WriteLine(ConsoleColor.Blue, para.Range.Text);
                    ConsoleC.WriteLine(ConsoleColor.Red, "This paragraph seems to contain a list. Consider rephrasing as a bulleted or numbered list.");
                }
            }
            ConsoleC.WriteLine(ConsoleColor.White, "The check for inline lists is complete.");
        }

        public static void RunTests()
        {
            string[] texts =
{
                "This is a paragraph that does not contain a list.",
                "This is a paragraph that contains a comma, another comma, and a third comma, but is not enumerated.",
                "This is a) a sentence, b) difficult to read, c) hard to maintain, d) bad practice, and e) difficult to read.",
                "This is A. a sentence, B. difficult to read, C. hard to maintain, D. bad practice, and E. difficult to read.",
                "This is (1) a sentence, (2) difficult to read, (3) hard to maintain, (4) bad practice, and (5) difficult to read.",
                "1) Don't pick up the phone, you know he's only calling coz he's drunk and alone. 2)Don't be his friend — you're only going to wake up in his bed in the morning."
            };

            foreach (string text in texts)
            {
                Console.WriteLine(IsMatch(text) + "    " + text);
            }
        }
    }
}
