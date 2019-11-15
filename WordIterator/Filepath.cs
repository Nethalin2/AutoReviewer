using System;
using System.IO;

public class Filepath
{
    private static readonly String folderDuncan = @"C:\Users\Duncan Ritchie\Documents\InformationCatalyst\AutoReviewer\AutoreviewerSideAssets";
    private static readonly String folderMark = @"C:\Users\netha\Documents\FSharpTest\FTEST";
    private static readonly String fileDuncan = "EU-ID D01 - ZDMP-ID D1.1 - Project Handbook - Annex - StyleGuide v1.0.2.docx";
    private static readonly String fileMark = "ftestdoc3.docx";
    private static readonly String fullDuncan = Path.Combine(folderDuncan, fileDuncan);
    private static readonly String fullMark = Path.Combine(folderMark, fileMark);
    private static readonly bool usingMarksFile = File.Exists(fullMark);
    private static readonly bool usingDuncansFile = File.Exists(fullDuncan);
    private static readonly FileNotFoundException exception = new FileNotFoundException("Neither Mark's nor Duncan's file was found"); 


    public static String Full()
    {
        if (usingMarksFile)
        {
            return fullMark;
        }
        else if (usingDuncansFile)
        {
            return fullDuncan;
        }
        else
        {
            throw exception;
        }
    }

    public static String Folder()
    {
        if (usingMarksFile)
        {
            return folderMark;
        }
        else if (usingDuncansFile)
        {
            return folderDuncan;
        }
        else
        {
            throw exception;
        }
    }

    public static String FileOnly()
    {
        if (usingMarksFile)
        {
            return fileMark;
        }
        else if (usingDuncansFile)
        {
            return fileDuncan;
        }
        else
        {
            throw exception;
        }
    }
}
