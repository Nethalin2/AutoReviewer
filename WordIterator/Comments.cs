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
            //Load the default instance of Document class.
            Document doc = LoadDocument.Default();

            //Get a paragraph to comment on.
            Paragraph para = doc.Paragraphs[1];

            // Add a comment.
            object text = "Add a comment to the first paragraph.";
            doc.Comments.Add(para.Range, ref text);

            //Save to a new file.
            doc.SaveAs2(Filepath.Full().Replace(".docx", "_2.docx"));
        }
    }
}
