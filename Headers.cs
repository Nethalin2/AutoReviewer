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
        InteropManager im;
        //Document wordDoc = im.getWord();
        public Headers()
        {
            im = new InteropManager("C:\\Users\\netha\\Documents\\FSharpTest\\FTEST", "justatest.docx");
        }

        public string ShortString(string InputText, int MaxLength)
        {
       
        int l = InputText.Length;
        int length = (l <= MaxLength ? l : MaxLength);
        string newString = InputText.Substring(0, length);
            return newString;
    }

    public void DetectHeader()
        {
          Word.Application wordDoc =  im.getWord(); 
          foreach (Paragraph paragraph in wordDoc.Application.ActiveDocument.Paragraphs)
          {
                Style style = paragraph.get_Style() as Style;
                int position = paragraph.ParaID;
                string styleName = style.NameLocal;
                string text = paragraph.Range.Text;
                Console.WriteLine(styleName + " " + position+"    "+(ShortString(text, 20).Replace("/n", " ").Replace("/r", " ").Trim()));
                
                //This checks the spacing after every paragraph.
                if (paragraph.Format.SpaceAfter == 6)
                {
                    Console.WriteLine("That's the correct spacing");
                } else
                {
                    Console.WriteLine("Spacing needs to change to 6pts");
                }

                //This checks the heading size.
                if (styleName == "Heading 1") 
                {
                    Console.WriteLine("Correct Heading Size");
                }
               else if (styleName == "Heading 2")
                {
                    Console.WriteLine("Correct Heading Size");
                }
                else if (styleName == "Heading 3")
                {
                    Console.WriteLine("Correct Heading Size");
                }
                else if (styleName == "Heading 4")
                {
                    Console.WriteLine("Correct Heading Size");
                }
                else if (styleName == "Heading 5")
                {
                    Console.WriteLine("Header is too small");
                }
                else if (styleName == "Heading 6")
                {
                    Console.WriteLine("Header is too small");      
                }
            }
        }
    }
}
