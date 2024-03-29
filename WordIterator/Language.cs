﻿using System;
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
    class Language
    {
        
        public static void LanguageChecker(Document doc)
        {
            try
            {
                int count = doc.Words.Count;

                int countUKEnglish = 0;
                int countUSEnglish = 0;
                int countNotUKUSEnglish = 0;

                ConsoleC.WriteLine(ConsoleColor.White, "Checking the language of every word...");

                for (int k = 1; k <= count; k++)
                {
                    //// Write a marker of where we are in the document every kth word.
                    if (k % 50 == 0)
                    {
                        ConsoleC.Write(ConsoleColor.Black, ConsoleColor.Gray, " " + k + " / " + count + " ");
                    }

                    string text = doc.Words[k].Text;

                    //// Check language

                    if (doc.Words[k].LanguageID == WdLanguageID.wdEnglishUK)
                    {
                        ConsoleC.Write(ConsoleColor.Green, text);
                        countUKEnglish++;
                        // ConsoleC.WriteLine(ConsoleColor.Green, "\nThis is a UK/US English word.");
                    }
                    else if (doc.Words[k].LanguageID == WdLanguageID.wdEnglishUS)
                    {
                        ConsoleC.Write(ConsoleColor.Yellow, text);
                        countUSEnglish++;
                        if (countUSEnglish % 10 == 1)
                        {
                            Comments.Add(doc, k, "This is US English but should be UK English.");
                        }
                    }
                    else
                    {
                        ConsoleC.WriteLine(ConsoleColor.Red, "\n" + text);
                        ConsoleC.WriteLine(ConsoleColor.Red, "This is not a UK or US English word.");
                        countNotUKUSEnglish++;
                        if (countNotUKUSEnglish % 10 == 1)
                        {
                            Comments.Add(doc, k, "This is not UK English but should be.");
                        }
                    }

                }

                //// Give feedback after all the words have been checked.
                ConsoleC.WriteLine(ConsoleColor.White, "\nFinished checking language.");

                ConsoleC.WriteLine(ConsoleColor.Green, countUKEnglish + " words were UK English. This is good!");
                if (countUSEnglish > 0)
                {
                    ConsoleC.WriteLine(ConsoleColor.Yellow, countUSEnglish + " words were US English. Please change these to UK English.");
                }
                if (countNotUKUSEnglish > 0)
                {
                    ConsoleC.WriteLine(ConsoleColor.Red, countNotUKUSEnglish + " words were neither. Please change these to UK English.");
                }
            }
            catch (Exception ex)
            {
                ConsoleC.WriteLine(ConsoleColor.Red, ex.ToString());

            }
        }

    }
}
