using System;
//using System.IO;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
//using System.Windows.Forms;

namespace WordIterator
{
    class Headers
    {

        // private static InteropManager im = new InteropManager(Filepath.Folder(), Filepath.FileOnly());

        // private static Word.Application wordDoc = im.getWord();
        // private static Document doc = wordDoc.Application.ActiveDocument;

        // Document wordDoc = im.getWord();
        //public Headers()
        //{
        //}

        public static void DetectLineSpacingAfterBullets(Document doc)
        {
            ConsoleC.WriteLine(ConsoleColor.White, "Checking every bullet for 6pt line-spacing between indentation levels...");

            int badSpacingCount = 0;
            int badSpacingFailCount = 0;

            // foreach (Paragraph paragraph in wordDoc.Application.ActiveDocument.Paragraphs)
            for (int i = 1; i < doc.Paragraphs.Count; i++)
            {
                Paragraph paragraph = doc.Paragraphs[i];
                Paragraph paragraph2 = doc.Paragraphs[i + 1];

                string listString = paragraph.Range.ListFormat.ListString;
                string listString2 = paragraph2.Range.ListFormat.ListString;

                if (listString != "" && listString2 != "" && paragraph.Format.LeftIndent != paragraph2.Format.LeftIndent)
                {
                    Style style = paragraph.get_Style() as Style;
                    string styleName = style.NameLocal;

                    if (styleName != "Heading 1" && styleName != "Heading 2" && styleName != "Heading 3" && styleName != "Heading 4")
                    {
                        if (paragraph.Format.SpaceAfter == 6)
                        {
                            //Console.WriteLine(paragraph.Range.Text);
                            //Console.WriteLine("That's the correct spacing");
                        }
                        else
                        {
                            ConsoleC.WriteLine(ConsoleColor.Blue, paragraph.Range.Text);
                            //Console.WriteLine("This paragraph's left indent is different to the next paragraph's left indent.");

                            ConsoleC.WriteLine(ConsoleColor.Yellow, "Detected line-spacing that should be 6pt but isn’t.");

                            badSpacingCount++;

                            try
                            {
                                paragraph.Format.SpaceAfter = 6;
                                ConsoleC.WriteLine(ConsoleColor.Green, "Spacing has been changed to 6pt.");
                                Comments.Add(doc, paragraph, "Line-spacing has been changed to 6pt.");
                            }
                            catch
                            {
                                ConsoleC.WriteLine(ConsoleColor.Red, "Failed to automatically change line-spacing to 6pt.");
                                Comments.Add(doc, paragraph, "Line-spacing needs to change to 6pt.");
                                badSpacingFailCount++;
                            }
                        }
                    }
                    else
                    {
                        //ConsoleC.WriteLine(ConsoleColor.Green, "This paragraph is a heading.");
                    }
                }
            }

            //// Give feedback having gone through the document.
            ConsoleC.WriteLine(ConsoleColor.White, "Finished checking every bullet.");
            ConsoleC.WriteLine(
                (badSpacingCount == 0 ? ConsoleColor.Green : ConsoleColor.Yellow), 
                "There were " + badSpacingCount + " instances where the spacing after a bullet " +
                "needed to be changed to 6pt before a bullet of a different indentation."
            );
            ConsoleC.WriteLine(
                (badSpacingFailCount == 0 ? ConsoleColor.Green : ConsoleColor.Red),
                "There are " + badSpacingFailCount + " instances where this could not be corrected automatically."
            );

            //// Save to a new file.
            doc.SaveAs2(Filepath.Full().Replace(".docx", "_2.docx"));
        }

        //// Assuming the 'k'th paragraph in Document 'doc' is a first-level heading.
        //// Every first-level heading should be followed by either a second-level heading,
        //// circa 3 or 4 lines of body text and then a second-level heading,
        //// or body text and no subheading.
        public static void DetectParaAfterHeadingOne(Document doc, int k) {
            ConsoleC.WriteLine(ConsoleColor.White, "Checking what’s after the heading...");
            try {
                Paragraph paragraph2 = doc.Paragraphs[k + 1];
                try {
                    Style style2 = paragraph2.get_Style() as Style;
                    string styleName2 = style2.NameLocal;

                    if (styleName2 == "Heading 2") {
                        ConsoleC.WriteLine(ConsoleColor.Green, "First-level heading is followed by a second-level heading — good!");
                    }
                    //// Assuming 75 characters per line.
                    else if (paragraph2.Range.Text.Length > 150 && paragraph2.Range.Text.Length < 375) {
                        ConsoleC.WriteLine(ConsoleColor.Green, "First-level heading is followed by circa 3 or 4 lines of " + styleName2 + " — good!");
                    }
                    else {
                        try {
                            Paragraph paragraph3 = doc.Paragraphs[k + 2];
                            Style style3 = paragraph3.get_Style() as Style;
                            string styleName3 = style3.NameLocal;

                            if (styleName3 == "Heading 2") {
                                ConsoleC.WriteLine(ConsoleColor.Blue, paragraph2.Range.Text);
                                string message = "After a first-level heading and before a second-level heading, there should be circa 3 or 4 lines of body text or none.";
                                ConsoleC.WriteLine(ConsoleColor.Red, message);
                                Comments.Add(doc, paragraph2, message);
                            }
                            else {
                                ConsoleC.WriteLine(ConsoleColor.Green, "First-level heading is followed by " + styleName2 + " and " + styleName3 + " — good!");
                            }
                        }
                        catch {
                            ConsoleC.WriteLine(ConsoleColor.Green, "First-level heading is followed by " + styleName2 + " — good!");
                        }
                    }
                }
                catch {
                    ConsoleC.WriteLine(ConsoleColor.Red, "Something failed with the style detection.");
                }
            }
            catch {
                ConsoleC.WriteLine(ConsoleColor.Red, "The last paragraph in the document is should not be a heading.");
                Comments.Add(doc, doc.Paragraphs[k], "The last paragraph in the document should not be a heading.");
            }
        }

        public static void DetectHeaders(Document doc) {
            try {
                for (int i = 1; i < doc.Paragraphs.Count; i++) {
                    Paragraph paragraph = doc.Paragraphs[i];
               
                    Style style = paragraph.get_Style() as Style;
                    int position = paragraph.ParaID;
                    string styleName = style.NameLocal;
                    string text = paragraph.Range.Text;

                    ConsoleC.Write(ConsoleColor.DarkCyan, styleName + "   ");
                    ConsoleC.Write(ConsoleColor.Cyan, position + "    ");
                    ConsoleC.WriteLine(ConsoleColor.Blue, StringMethods.TrimAndShorten(text, 20));
                    ////This checks the spacing after every paragraph.
                    //if (position == 360681186)
                    //{
                    //    ConsoleC.WriteLine(ConsoleColor.Cyan, "Left indent: "+paragraph.Format.LeftIndent);
                    //}

                    //This checks the heading size.
                    if (styleName == "Heading 1") 
                    {
                        ConsoleC.WriteLine(ConsoleColor.Green, "Acceptable heading size.");

                        //// Check whether the first-level heading is followed by
                        //// either a second-level heading or a 3–4 lines of body text.
                        DetectParaAfterHeadingOne(doc, i);

                    }
                    else if (styleName == "Heading 2")
                    {
                        ConsoleC.WriteLine(ConsoleColor.Green, "Acceptable heading size.");
                    }
                    else if (styleName == "Heading 3")
                    {
                        ConsoleC.WriteLine(ConsoleColor.Green, "Acceptable heading size.");
                    }
                    else if (styleName == "Heading 4")
                    {
                        ConsoleC.WriteLine(ConsoleColor.Green, "Acceptable heading size.");
                    }
                    else if (styleName == "Heading 5")
                    {
                        ConsoleC.WriteLine(ConsoleColor.Red, "Header is too small.");
                        Comments.Add(doc, paragraph, "Heading level should be higher than 5.");
                    }
                    else if (styleName == "Heading 6")
                    {
                        ConsoleC.WriteLine(ConsoleColor.Red, "Header is too small.");
                        Comments.Add(doc, paragraph, "Heading level should be higher than 5, but is 6.");
                    }
                }
            }
            catch (Exception ex)
            {
                ConsoleC.WriteLine(ConsoleColor.Red, "Error in DetectHeaders() — " + ex.ToString());
            }
        }
    }
}
