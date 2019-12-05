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
        public static Document Default()
        {
            try
            {
                return AnyDoc(Filepath.Full());
            }
            catch
            {
                throw new Exception("Error loading default document.");
            }
        }
        public static Document AnyDoc(string filepath)
        {
            try
            {
                object fileName = filepath;

                Application wordApp = new Application { Visible = true };

                Document aDoc = wordApp.Documents.Open(ref fileName, ReadOnly: false, Visible: true);

                aDoc.Activate();

                return (aDoc);
            }
            catch
            {
                throw new Exception("Error loading document " + filepath + "!");
            }
        }
    }
}
