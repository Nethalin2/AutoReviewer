using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;

namespace WordIterator
{
    class InteropManager
    {
        public object fileName;
        public Word.Application wordApp;
        public Word.Document aDoc;

        public InteropManager(string path, string document)
        {
            Object wordObject = null;
            
            
            fileName = Path.Combine(path, document);
            wordApp = new Word.Application { Visible = true };
            aDoc = wordApp.Documents.Open(ref fileName, ReadOnly: false, Visible: true);
            aDoc.Activate();
            wordObject = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
        }

        public Word.Application getWord()
        {
            return wordApp;
        }
        public void getDocument()
        { 
        }

    }
}
