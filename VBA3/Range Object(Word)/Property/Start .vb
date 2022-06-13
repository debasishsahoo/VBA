pos = ActiveDocument.Paragraphs(2).Range.Start 
pos2 = ActiveDocument.Paragraphs(4).Range.End 
Set myRange = ActiveDocument.Range(Start:=pos, End:=pos2)

Set myRange = Selection.Range 
myRange.SetRange Start:=myRange.Start + 1, End:=myRange.End