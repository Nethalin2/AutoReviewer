using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordIterator
{
    class LoadDocument
    {
        public static Word.Document Default()
        {
            try
            {
                ConsoleC.WriteLine(ConsoleColor.White, "Trying to load a file...");
               
                object fileName = Filepath.Full();

                Application wordApp = new Word.Application { Visible = true };

                Document aDoc = wordApp.Documents.Open(ref fileName, ReadOnly: false, Visible: true);

                aDoc.Activate();
                Application word = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");

                word.Visible = true;
                word.ScreenUpdating = false;

                ConsoleC.WriteLine(ConsoleColor.Green, "The file has loaded.");

                return word.ActiveDocument;
            }
            catch
            {
                throw new Exception("Error loading default document!");
            }
        }    
    }
}
