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
        public static void AddComment()
        {
            Console.WriteLine("Trying to write a comment on the first paragraph!");

            //Load the default instance of Document class.
            Document doc = LoadDocument.Default();

            Console.ForegroundColor = ConsoleColor.Green;

            for (int i = 1; i < doc.Paragraphs.Count; i++)
            {
                // Add a comment.
                try
                {
                    object text = "This is a comment on Paragraph "+i+".";
                    doc.Comments.Add(doc.Paragraphs[i].Range, ref text);
                    Console.WriteLine("Added a comment on paragraph " + i + "!");
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Failed to add a comment to paragraph " + i + " — "+ex.ToString());
                    Console.ForegroundColor = ConsoleColor.Green;
                }
            }

            /*
            //Get a paragraph to comment on.
            Paragraph para = doc.Paragraphs[1];

            // Add a comment.
            object text = "Add a comment to the first paragraph.";
            doc.Comments.Add(para.Range, ref text);
            */

            //Save to a new file.
            doc.SaveAs2(Filepath.Full().Replace(".docx", "_2.docx"));
        }
    }
}
