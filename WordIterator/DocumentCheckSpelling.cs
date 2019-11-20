using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace WordIterator
{
    class DocumentCheckSpelling
    {
        public static void DocCheckSpelling(Document doc)
        {
            // var app = new Microsoft.Office.Interop.Word.Application();

            try
            {
                // Document doc = app.Documents.Open("C:\\Users\\netha\\Documents\\FSharpTest\\FTEST\\justatest.docx");
                //Document doc = LoadDocument.Default();

                foreach (var word in doc.Words.Cast<Range>())
                {
                    if (word.SpellingErrors.Count > 0)
                    {
                        Console.WriteLine("This word is spelt incorrectly " + word.Text);
                    }
                }
            }
            catch
            {
                //Use Try/Catch to avoid persisting Word processes in the event of an exception
            }
            finally
            {
                // Console.ReadLine();
                // app.Quit();
            }
        }
    }
}
