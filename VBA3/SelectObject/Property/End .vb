pos = Selection.End 
Set myRange = ActiveDocument.Range(Start:=pos, End:=pos) 
ActiveDocument.Fields.Add Range:=myRange, Type:=wdFieldAuthor