using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Windows;



namespace WordIterator
{
    class Program
    {
        
        static void Main(string[] args)
        {   
            

            //// Load a document we can play with.
            Document doc = LoadDocument.Default();
            DocumentCheckSpelling.DocCheckSpelling(doc);
            
            InlineLists.DetectAll(doc);

            Headers.DetectHeaders(doc);
            Headers.DetectLineSpacingAfterBullets(doc);

            Language.LanguageChecker(doc);
       
            //// Save to a new file.
            doc.SaveAs2(Filepath.Full().Replace(".docx", "_2.docx"));
            doc.Close();
            Document doc2 = LoadDocument.AnyDoc(@"C:\Users\netha\Documents\FSharpTest\FTEST\ftestdoc3_2.docx");

            //// Keep the console open even when the program has finished.
            Word.Application wordApp = new Word.Application { Visible = true };
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            for (int i = 1; i <= doc2.Comments.Count; i++)
            {

               
                Console.WriteLine(doc2.Comments[i].Author);
                Console.WriteLine(doc2.Comments[i].Range.Text);
                Console.WriteLine("The number of comments is: " + doc2.Comments.Count);
                var regex = new System.Text.RegularExpressions.Regex("Ignore", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                Console.WriteLine("Regex result is " + regex.IsMatch(doc2.Comments[i].Range.Text).ToString());
                Console.WriteLine("Comment Content is : '" + doc2.Comments[i].Range.Text+"'");

                if (regex.IsMatch(doc2.Comments[i].Range.Text))
                {
                    Console.WriteLine("Deleting a comment that reads - '" + doc2.Comments[i].Range.Text + "'");
                    //deletes comment text
                    //doc2.Comments[i].Range.Delete();

                    //doesnt work
                    //doc2.Comments[i].Scope.Delete();

                    //deletes text author and entire comment but misses some comments
                    doc2.Comments[i].DeleteRecursively();
                }

            }

            

            ConsoleC.WriteLine(ConsoleColor.Green, "\nThe program has finished.");
            Console.ReadLine();
            
        }

       
    }
   
   
}



