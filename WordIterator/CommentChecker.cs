using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace WordIterator
{
    class CommentChecker
    {
        public static void CommCheck(Document doc)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            //Document doc = LoadDocument.AnyDoc(@"C:\Users\netha\Documents\FSharpTest\FTEST\ftestdoc3_2.docx");
            for (int i = 1; i <= doc.Comments.Count; i++)
            {


                Console.WriteLine(doc.Comments[i].Author);
                Console.WriteLine(doc.Comments[i].Range.Text);
                Console.WriteLine("The number of comments is: " + doc.Comments.Count);
                var regex = new System.Text.RegularExpressions.Regex("Ignore", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                Console.WriteLine("Regex result is " + regex.IsMatch(doc.Comments[i].Range.Text).ToString());
                Console.WriteLine("Comment Content is : '" + doc.Comments[i].Range.Text + "'");

                if (regex.IsMatch(doc.Comments[i].Range.Text))
                {
                    Console.WriteLine("Deleting a comment that reads - '" + doc.Comments[i].Range.Text + "'");
                    //deletes comment text
                    //doc2.Comments[i].Range.Delete();

                    //doesnt work
                    //doc2.Comments[i].Scope.Delete();

                    //deletes text author and entire comment but misses some comments
                    doc.Comments[i].DeleteRecursively();
                }

            }

        }
    }
}
