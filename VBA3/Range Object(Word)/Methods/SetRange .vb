Set myRange = ActiveDocument.Paragraphs(1).Range 
myRange.SetRange Start:=myRange.Start, _ 
End:=ActiveDocument.Paragraphs(3).Range.End


Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
myRange.InsertAfter "Hello " 
myRange.SetRange Start:=myRange.Start, End:=Selection.End