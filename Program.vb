For k = 1 To count
    Dim text = document.Words(k).Text
    Dim Bold = document.Words(k).Bold
    Dim ActiveDocument As [With]
    [If].LanguageDetected = [True]
    Dim x As [Then] = MsgBox("This document has already ", __ And "been checked. Do you want to check ", __ And "it again?", vbYesNo)
    Dim x As [If] = vbYes
    [Then].LanguageDetected = [False].DetectLanguage
    Dim [If] As [End]
    Dim [End] As [Else].DetectLanguage
    Dim _ As [If]
    [If].Range.LanguageID = wdEnglishUS
    Dim MsgBox As [Then]
    "This is a U.S. English document."
    Dim MsgBox As [Else]
    "This is not a U.S. English document."
    Dim [If] As [End]
    Dim [With] As [End]
    Dim SpellingChecked = document.Words(k).SpellingChecked
Next