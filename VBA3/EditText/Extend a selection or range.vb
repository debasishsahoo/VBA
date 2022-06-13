Sub ExtendSelection() 
 Selection.MoveEnd Unit:=wdWord, Count:=3 
End Sub

Sub ExtendRange() 
 Dim rngParagraphs As Range 
 
 Set rngParagraphs = ActiveDocument.Paragraphs(1).Range 
 rngParagraphs.MoveEnd Unit:=wdParagraph, Count:=2 
End Sub