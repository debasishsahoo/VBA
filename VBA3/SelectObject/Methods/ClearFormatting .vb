Sub ClrFmtg() 
 ActiveDocument.Select 
 Selection.ClearFormatting 
End Sub

Sub ClrFmtg2() 
 ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(2).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End).Select 
 Selection.ClearFormatting 
End Sub