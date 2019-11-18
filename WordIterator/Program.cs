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

            // Comments.AddToEveryPara();

            LanguageChecker();

            //Headers thisHeader = new Headers();
            //// thisHeader.DetectHeader();
            //thisHeader.DetectLineSpacingAfterBullets();
            Console.ReadLine();

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

                int countUKEnglish = 0;
                int countUSEnglish = 0;
                int countNotUKUSEnglish = 0;

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Checking the language of every word...");

                for (int k = 1; k <= count; k++)
                {
                    //// Write a marker of where we are in the document every kth word.
                    if (k % 50 == 0)
                    {
                        Console.BackgroundColor = ConsoleColor.Gray;
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.Write(" "+k+" / "+count+" ");
                        Console.BackgroundColor = ConsoleColor.Black;
                    }

                    string text = document.Words[k].Text;
                    //int Bold = document.Words[k].Bold;

                    //for (k = 1; (k <= count); k++)
                    //{


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
           

                    //try
                    //{
                    //    document.Words[k].LanguageID = WdLanguageID.wdEnglishUK;
                    //}
                    //catch
                    //{
                    //    Console.WriteLine("Correcting language failed!");
                    //}

                    //// Check language

                    if (document.Words[k].LanguageID == WdLanguageID.wdEnglishUK)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write(text);
                        countUKEnglish++;
                        // Console.WriteLine("This is a UK/US English word.");
                    }
                    else if (document.Words[k].LanguageID == WdLanguageID.wdEnglishUS)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.Write(text);
                        countUSEnglish++;
                        if (countUSEnglish % 10 == 1)
                        {
                            Comments.Add(aDoc, k, "This is US English but should be UK English.");
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("\n" + text);
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("This is not a UK or US English word.");
                        countNotUKUSEnglish++;
                        if (countNotUKUSEnglish % 10 == 1)
                        {
                            Comments.Add(aDoc, k, "This is not UK English but should be.");
                        }
                        //try
                        //{
                        //    document.Words[k].LanguageID = WdLanguageID.wdEnglishUK;
                        //}
                        //catch
                        //{
                        //    Console.WriteLine("Correcting language failed!");
                        //}
                    }

                    //// Check whether spellcheck is checked.

                    /*
                    bool SpellingChecked = document.Words[k].SpellingChecked;
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Spelling check is set to " + (SpellingChecked ? "true" : "false"));
                    //object SpellingChecked = document.Words(k).SpellingChecked;
                    */

                    //// Save to a new file.
                    try
                    {
                        aDoc.SaveAs2(Filepath.Full().Replace(".docx", "_2.docx"));
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Failed to save new file — "+ex.ToString());
                    }

                }
                
                //// Give feedback after all the words have been checked.
                
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("\nFinished checking language.");

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(countUKEnglish + " words were UK English. This is good!");
                if (countUSEnglish > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(countUSEnglish + " words were US English. Please change these to UK English.");
                }
                if (countNotUKUSEnglish > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(countNotUKUSEnglish + " words were neither. Please change these to UK English.");
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

                //Console.ForegroundColor = ConsoleColor.Green;
                //Console.WriteLine("Finished iterating across document.");

                //if (document.Paragraphs.Count > 0)
                //{
                //    var paragraph = document.Paragraphs.First;
                //    var lastCharPos = paragraph.Range.Sentences.First.End - 1;
                //    MessageBox.Show(lastCharPos.ToString());
                //}
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());

            }
            Console.ResetColor();
            // Console.ReadLine();
            if (word != null)
            {
                word.Quit();
            }

        }
    }
   
}
