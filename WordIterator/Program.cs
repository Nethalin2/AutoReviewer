﻿using System;
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
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("\n"+text);
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("This is not a UK or US English word.");
                            countNotUKUSEnglish++;
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
