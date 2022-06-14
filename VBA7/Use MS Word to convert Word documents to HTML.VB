Option Explicit

Doc2HTML "C:\Documents and Settings\MyUserID\My Documents\resume.doc"

Sub Doc2HTML( myFile )
' This subroutine opens a Word document,
' then saves it as HTML, and closes Word.
' If the HTML file exists, it is overwritten.
' If Word was already active, the subroutine
' will leave the other document(s) alone and
' close only its "own" document.
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
    ' Standard housekeeping
    Dim objDoc, objFile, objFSO, objWord, strFile, strHTML

    Const wdFormatDocument                    =  0
    Const wdFormatDocument97                  =  0
    Const wdFormatDocumentDefault             = 16
    Const wdFormatDOSText                     =  4
    Const wdFormatDOSTextLineBreaks           =  5
    Const wdFormatEncodedText                 =  7
    Const wdFormatFilteredHTML                = 10
    Const wdFormatFlatXML                     = 19
    Const wdFormatFlatXMLMacroEnabled         = 20
    Const wdFormatFlatXMLTemplate             = 21
    Const wdFormatFlatXMLTemplateMacroEnabled = 22
    Const wdFormatHTML                        =  8
    Const wdFormatPDF                         = 17
    Const wdFormatRTF                         =  6
    Const wdFormatTemplate                    =  1
    Const wdFormatTemplate97                  =  1
    Const wdFormatText                        =  2
    Const wdFormatTextLineBreaks              =  3
    Const wdFormatUnicodeText                 =  7
    Const wdFormatWebArchive                  =  9
    Const wdFormatXML                         = 11
    Const wdFormatXMLDocument                 = 12
    Const wdFormatXMLDocumentMacroEnabled     = 13
    Const wdFormatXMLTemplate                 = 14
    Const wdFormatXMLTemplateMacroEnabled     = 15
    Const wdFormatXPS                         = 18
    Const wdFormatOfficeDocumentTemplate      = 23 (1)
    Const wdFormatMediaWiki                   = 24(1) (2)

    ' Create a File System object
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ' Create a Word object
    Set objWord = CreateObject( "Word.Application" )

    With objWord
        ' True: make Word visible; False: invisible
        .Visible = True

        ' Check if the Word document exists
        If objFSO.FileExists( myFile ) Then
            Set objFile = objFSO.GetFile( myFile )
            strFile = objFile.Path
        Else
            WScript.Echo "FILE OPEN ERROR: The file does not exist" & vbCrLf
            ' Close Word
            .Quit
            Exit Sub
        End If
        ' Build the fully qualified HTML file name
        strHTML = objFSO.BuildPath( objFile.ParentFolder, _
                  objFSO.GetBaseName( objFile ) & ".html" )

        ' Open the Word document
        .Documents.Open strFile

        ' Make the opened file the active document
        Set objDoc = .ActiveDocument

        ' Save as HTML
        objDoc.SaveAs strHTML, wdFormatFilteredHTML

        ' Close the active document
        objDoc.Close

        ' Close Word
        .Quit
    End With
End Sub