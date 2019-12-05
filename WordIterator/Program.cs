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
            CommentChecker.CommCheck(doc2);
            //// Keep the console open even when the program has finished.
            Word.Application wordApp = new Word.Application { Visible = true };
           

            
            
            ConsoleC.WriteLine(ConsoleColor.Green, "\nThe program has finished.");
            Console.ReadLine();
            
        }

       
    }
   
   
}



