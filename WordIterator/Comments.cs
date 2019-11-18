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
            //Load the default instance of Document class

            Document doc = LoadDocument.Default();

            //Get the paragraph that you want to add comment

            Paragraph para = doc.Sections[0].Paragraphs[6];


            //Insert a start comment mark at the beginning of the paragraph

            CommentMark startCommentMark = new CommentMark(doc);

            startCommentMark.Type = CommentMarkType.CommentStart;

            para.ChildObjects.Insert(0, startCommentMark);


            //Add an end comment mark at the end of the paragraph

            CommentMark endCommentMark = new CommentMark(doc);

            endCommentMark.Type = CommentMarkType.CommentEnd;

            para.ChildObjects.Add(endCommentMark);


            //Add a comment to the paragraph and specify the content and author

            Comment comment = new Comment(doc);

            comment.Body.AddParagraph().AppendText("This paragraph explains how to write an abstract.");

            comment.Format.Author = "John";

            para.ChildObjects.Add(comment);


            //Save to file

            doc.SaveToFile("CommentOnPara.docx", FileFormat.Docx2013);
        }
    }
}
