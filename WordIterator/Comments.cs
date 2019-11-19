using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordIterator
{
    class Comments
    {
        //// Add() is overloaded such that it will accept either a Document, Paragraph, and String,
        //// or a Document, int, and string. The string is coerced into an object to work with interop.
        //// If the second parameter is an int 'k', it is treated as referring to the 'k'th word in the Document.
        public static void Add(Document doc, Paragraph placeForComment, object comment)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("\nAdding a comment — ");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(comment);

                doc.Comments.Add(placeForComment.Range, ref comment);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to add a comment to paragraph!");
            }
        }

        public static void Add(Document doc, int k, object comment)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("\nAdding a comment of ");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.Write(comment);
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(" to word #"+k+".");

                doc.Comments.Add(doc.Words[k], ref comment);
            }
            
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to add a comment to word #" + k + "!");
            }
        }

        //// This function demonstrates that we can add a comment to (nearly) every paragraph if we like.
        public static void AddToEveryPara(Document doc)
        {
            Console.WriteLine("Trying to write a comment on all the paragraphs!");

            //// Load the default instance of Document class.
            // Document doc = LoadDocument.Default();

            Console.ForegroundColor = ConsoleColor.Green;

            for (int i = 1; i < doc.Paragraphs.Count; i++)
            {
                // Add a comment.
                try
                {
                    object text = "This is a comment on Paragraph "+i+".";
                    doc.Comments.Add(doc.Paragraphs[i].Range, ref text);
                    Console.WriteLine("Added a comment on Paragraph " + i + "!");
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Failed to add a comment to Paragraph " + i + " — "+ex.ToString());
                    Console.ForegroundColor = ConsoleColor.Green;
                }
            }

            //Save to a new file.
            doc.SaveAs2(Filepath.Full().Replace(".docx", "_2.docx"));
        }
    }
}
