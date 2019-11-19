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
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Checking every bullet for 6pt line-spacing between indentation levels...");

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
                            Console.ForegroundColor = ConsoleColor.Blue;
                            Console.WriteLine(paragraph.Range.Text);
                            //Console.WriteLine("This paragraph's left indent is different to the next paragraph's left indent.");

                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine("Detected line-spacing that should be 6pt but isn’t.");

                            badSpacingCount++;

                            try
                            {
                                paragraph.Format.SpaceAfter = 6;
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("Spacing has been changed to 6pt.");
                                Comments.Add(doc, paragraph, "Line-spacing has been changed to 6pt.");
                            }
                            catch
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Failed to automatically change line-spacing to 6pt.");
                                Comments.Add(doc, paragraph, "Line-spacing needs to change to 6pt.");
                                badSpacingFailCount++;
                            }
                        }
                    }
                    else
                    {
                        //Console.WriteLine("This paragraph is a heading.");
                    }
                }
            }

            //// Give feedback having gone through the document.
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Finished checking every bullet.");
            Console.ForegroundColor = badSpacingCount == 0 ? ConsoleColor.Green : ConsoleColor.Yellow;
            Console.WriteLine("There were " + badSpacingCount + " instances where the spacing after a bullet needed to be changed to 6pt before a bullet of a different indentation.");
            Console.ForegroundColor = badSpacingFailCount == 0 ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine("There are " + badSpacingFailCount + " instances where this could not be corrected automatically.");

            //// Save to a new file.
            doc.SaveAs2(Filepath.Full().Replace(".docx", "_2.docx"));
        }

        public static void DetectHeader(Document doc)
        {
            for (int i = 1; i < doc.Paragraphs.Count; i++)
            {
                Paragraph paragraph = doc.Paragraphs[i];
                //Paragraph paragraph2 = doc.Paragraphs[i + 1];
               
               
                Style style = paragraph.get_Style() as Style;
                int position = paragraph.ParaID;
                string styleName = style.NameLocal;
                string text = paragraph.Range.Text;

                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.Write(styleName + "   ");
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write(position + "    ");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(StringMethods.TrimAndShorten(text, 20));
                //This checks the spacing after every paragraph.
                //if (position == 360681186)
                //{

                //Console.ForegroundColor = ConsoleColor.Cyan;

                //Console.WriteLine("Left indent: "+paragraph.Format.LeftIndent);
                //}

                Console.ForegroundColor = ConsoleColor.Green;

                //This checks the heading size.
                if (styleName == "Heading 1") 
                {
                    Console.WriteLine("Acceptable heading size.");
                }
                else if (styleName == "Heading 2")
                {
                    Console.WriteLine("Acceptable heading size.");
                }
                else if (styleName == "Heading 3")
                {
                    Console.WriteLine("Acceptable heading size.");
                }
                else if (styleName == "Heading 4")
                {
                    Console.WriteLine("Acceptable heading size.");
                }
                else if (styleName == "Heading 5")
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Header is too small.");
                    Comments.Add(doc, paragraph, "Heading level should be higher than 5.");
                }
                else if (styleName == "Heading 6")
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Header is too small.");
                    Comments.Add(doc, paragraph, "Heading level should be higher than 5, but is 6.");
                }
            }
        }
    }
}
