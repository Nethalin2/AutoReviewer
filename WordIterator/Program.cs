using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Windows;

namespace WordIterator
{
    class Program
    {
        static void Main(string[] args)
        {
            LanguageChecker();
            //DocumentCheckSpelling.DocCheckSpelling();

            //Headers thisHeader = new Headers();
            //thisHeader.DetectHeader();
            //Console.ReadLine();

        }
        public static void LanguageChecker()
        {
            Object wordObject = null;
            Word.Application word = null;
            Document document = null;
            Word.Document aDoc = null;

            try
            {
               
                object fileName = Path.Combine("C:\\Users\\netha\\Documents\\FSharpTest\\FTEST", "justatest.docx");

                Word.Application wordApp = new Word.Application { Visible = true };

                aDoc = wordApp.Documents.Open(ref fileName, ReadOnly: false, Visible: true);

                aDoc.Activate();
                wordObject = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");

                word = (Word.Application)wordObject;
                word.Visible = false;
                word.ScreenUpdating = false;
                string fullPath = word.ActiveDocument.FullName;

                document = word.ActiveDocument;
                int count = document.Words.Count;



               

                for (int k = 1; (k <= count); k++)
                {
                    string text = document.Words[k].Text;
                    int Bold = document.Words[k].Bold;



                    if (word.ActiveDocument.LanguageDetected == true)
                    {



                        word.ActiveDocument.LanguageDetected = false;
                        word.ActiveDocument.DetectLanguage();



                    }
                    else
                    {
                        word.ActiveDocument.DetectLanguage();
                    }
                    if (word.ActiveDocument.Words[k].LanguageID == WdLanguageID.wdEnglishUS || word.ActiveDocument.Words[k].LanguageID == WdLanguageID.wdEnglishUK)
                    {

                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.WriteLine("This is an English document.");
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("This is not an English word.");
                        document.Words[k].Font.ColorIndex = Word.WdColorIndex.wdYellow;
                        Console.WriteLine(document.Words[k].Text);
                    }

                }
                }
                
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
            finally { 
                aDoc.Save();
                aDoc.Close();
                word.Quit();
                Console.ReadLine();
            }


        }

    }
   
   
}



