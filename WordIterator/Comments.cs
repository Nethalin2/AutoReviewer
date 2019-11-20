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
                ConsoleC.Write(ConsoleColor.White, "\nAdding a comment — ");
                ConsoleC.WriteLine(ConsoleColor.Blue, comment);

                doc.Comments.Add(placeForComment.Range, ref comment);
            }
            catch
            {
                ConsoleC.WriteLine(ConsoleColor.Red, "Failed to add a comment to paragraph!");
            }
        }

        public static void Add(Document doc, int k, object comment)
        {
            try
            {
                ConsoleC.Write(ConsoleColor.White, "\nAdding a comment of ");
                ConsoleC.Write(ConsoleColor.Blue, comment);
                ConsoleC.WriteLine(ConsoleColor.White, " to word #"+k+".");

                doc.Comments.Add(doc.Words[k], ref comment);
            }
            
            catch
            {
                ConsoleC.WriteLine(ConsoleColor.Red, "Failed to add a comment to word #" + k + "!");
            }
        }

        //// This function demonstrates that we can add a comment to (nearly) every paragraph if we like.
        public static void AddToEveryPara(Document doc)
        {
            ConsoleC.WriteLine(ConsoleColor.White, "Trying to write a comment on all the paragraphs!");

            //// Load the default instance of Document class.
            // Document doc = LoadDocument.Default();

            for (int i = 1; i < doc.Paragraphs.Count; i++)
            {
                // Add a comment.
                try
                {
                    object text = "This is a comment on Paragraph "+i+".";
                    doc.Comments.Add(doc.Paragraphs[i].Range, ref text);
                    ConsoleC.WriteLine(ConsoleColor.Green, "Added a comment on Paragraph " + i + "!");
                }
                catch (Exception ex)
                {
                    ConsoleC.WriteLine(ConsoleColor.Red, "Failed to add a comment to Paragraph " + i + " — "+ex.ToString());
                }
            }

            //Save to a new file.
            doc.SaveAs2(Filepath.Full().Replace(".docx", "_2.docx"));
        }
    }
}
