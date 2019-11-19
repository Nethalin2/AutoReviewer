using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

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
    }
}
