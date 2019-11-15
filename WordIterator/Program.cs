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
    class Program
    {
        static void Main(string[] args)
        {
            LanguageChecker();

            //Headers thisHeader = new Headers();
            //// thisHeader.DetectHeader();
            //thisHeader.DetectLineSpacingAfterBullets();
            //Console.ReadLine();

        }
        public static void LanguageChecker()
        {
            Object wordObject = null;
            Word.Application word = null;
            Document document = null;


            try
            {
                object fileName = Filepath.Full(); 
                // object fileName = Path.Combine(@"C:\Users\Duncan Ritchie\Documents\InformationCatalyst\AutoReviewer\AutoreviewerSideAssets", "EU-ID D01 - ZDMP-ID D1.1 - Project Handbook - Annex - StyleGuide v1.0.2.docx");
                //}
                //catch
                //{
                //    throw new FileNotFoundException("Filepath.Full() not working");
                //}
              
                Word.Application wordApp = new Word.Application { Visible = true };

                Document aDoc = wordApp.Documents.Open(ref fileName, ReadOnly: false, Visible: true);

                aDoc.Activate();
                wordObject = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");

                word = (Microsoft.Office.Interop.Word.Application)wordObject;
                word.Visible = false;
                word.ScreenUpdating = false;
                string fullPath = word.ActiveDocument.FullName;

                document = word.ActiveDocument;

                int count = document.Words.Count;
                for (int k = 1; k <= count; k++)
                {
                    string text = document.Words[k].Text;
                    // int Bold = document.Words[k].Bold;



                        //if (document.LanguageDetected == true)
                        //{
                        //    document.LanguageDetected = false;
                        //    document.DetectLanguage();

                        //}
                        //else
                        //{
                        //    try
                        //    {
                        //        document.DetectLanguage();

                        //    }
                        //    catch
                        //    {
                        //        // Console.WriteLine("DetectLanguage() failed!");
                        //    }

                        //}
                        if (document.Words[k].LanguageID == WdLanguageID.wdEnglishUK || document.Words[k].LanguageID == WdLanguageID.wdEnglishUS)
                        {
                            // Console.WriteLine("This is a UK/US English word.");
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.DarkYellow;
                            Console.WriteLine(text);
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine("This is not a UK or US English word.");
                            //try
                            //{
                            //    document.Words[k].LanguageID = WdLanguageID.wdEnglishUK;
                            //}
                            //catch
                            //{
                            //    Console.WriteLine("Correcting language failed!");
                            //}
                        }

                        //object SpellingChecked = document.Words(k).SpellingChecked;
                    }
                    //Boolean SpellingChecked = document.Words[k].SpellingChecked;
                    //Console.WriteLine(text + (SpellingChecked ? "true" : "false"));
                    //MessageBox.Show(text + " " + Bold.ToString());
                    //MessageBox.Show(text)
                    //Console.WriteLine(text + " " + Bold.ToString());
  
                //for (int r = 0; r <= document.Characters.Count; r++)
                //{
                //    Console.WriteLine(document.Characters[r]+ " " + document.Characters[r].CharacterStyle.toString());
                //}


                //if (document.Paragraphs.Count > 0)
                //{
                //    var paragraph = document.Paragraphs.First;
                //    var lastCharPos = paragraph.Range.Sentences.First.End - 1;
                //    MessageBox.Show(lastCharPos.ToString());
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            //Console.ReadLine();
            //if (word != null)
            //{
            //     word.Quit();
            //}
           

        }
    }
   
}
