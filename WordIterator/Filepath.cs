using System;
using System.IO;

public class Filepath
{
    private static readonly String folderDuncan = @"C:\Users\Duncan Ritchie\Documents\InformationCatalyst\AutoReviewer\AutoreviewerSideAssets";
    private static readonly String folderMark = @"C:\Users\netha\Documents\FSharpTest\FTEST";
    private static readonly String fileDuncan = "EU-ID D01 - ZDMP-ID D1.1 - Project Handbook - Annex - StyleGuide v1.0.2.docx";
    // private static readonly String fileDuncan = "windows.docx";
    // private static readonly String fileDuncan = "handbookAllUKEnglish.docx";
    private static readonly String fileMark = "ftestdoc3.docx";
    private static readonly String fullDuncan = Path.Combine(folderDuncan, fileDuncan);
    private static readonly String fullMark = Path.Combine(folderMark, fileMark);
    private static readonly bool usingMarksFile = File.Exists(fullMark);
    private static readonly bool usingDuncansFile = File.Exists(fullDuncan);
    private static readonly String newFileEnding = "_2";
    private static readonly FileNotFoundException exception = new FileNotFoundException("Neither Mark’s nor Duncan’s file was found");

    //// Used inside the three following functions (Full, Folder, File) to save repetition of the if-else statements.
    private static String IfStatements(String mark, String duncan)
    {
        if (usingMarksFile)
        {
            return mark;
        }
        else if (usingDuncansFile)
        {
            return duncan;
        }
        else
        {
            throw exception;
        }
    }

    public static String Full()
    {
        return IfStatements(fullMark, fullDuncan);
    }

    public static String Folder()
    {
        return IfStatements(folderMark, folderDuncan);
    }

    public static String FileOnly()
    {
        return IfStatements(fileMark, fileDuncan);
    }

    public static String FullNew()
    {
        return Full().Replace(".docx", newFileEnding+".docx");
    }
}
