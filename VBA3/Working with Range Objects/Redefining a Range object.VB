Sub ExpandRange() 
 Dim rngParagraph As Range 
 Set rngParagraph = Selection.Range 
 rngParagraph.SetRange Start:=rngParagraph.Start, _ 
 End:=rngParagraph.End + 10 
End Sub


Sub SelectRange() 
 Dim rngParagraph As Range 
 Set rngParagraph = ActiveDocument.Paragraphs(2).Range 
 rngParagraph.SetRange Start:=rngParagraph.Start, _ 
 End:=ActiveDocument.Paragraphs(3).Range.End 
 rngParagraph.Select 
 Selection.Font.Italic = True 
End Sub