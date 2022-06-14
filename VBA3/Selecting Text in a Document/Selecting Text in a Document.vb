Sub SelectTable() 
 ActiveDocument.Tables(1).Select 
End Sub

Sub SelectField() 
 ActiveDocument.Fields(1).Select 
End Sub

Sub SelectRange() 
 Dim rngParagraphs As Range 
 Set rngParagraphs = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(1).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End) 
 rngParagraphs.Select 
End Sub