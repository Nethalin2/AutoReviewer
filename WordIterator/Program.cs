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

            // DocumentCheckSpelling.DocCheckSpelling(doc);

            // InlineLists.RunTests();

            Console.WriteLine("\n");
            

            InlineLists.DetectAll(doc);

            Console.WriteLine("\n");

            // Comments.AddToEveryPara(doc);

            Headers.DetectHeaders(doc);

            Console.WriteLine("\n");

            Headers.DetectLineSpacingAfterBullets(doc);

            // Language.LanguageChecker(doc);

            //// Save to a new file.
            string newFilepath = Filepath.FullNew();
            doc.SaveAs2(newFilepath);
            ConsoleC.WriteLine(ConsoleColor.Green, "\nA file has been created with comments and edits at "+newFilepath);

            //// Keep the console open even when the program has finished.
            ConsoleC.WriteLine(ConsoleColor.Green, "\nThe program has finished.");
            Console.ReadLine();
        }

       
    }
   
   
}



