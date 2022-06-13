Set myRange = Selection.Paragraphs(1).Range 
leng = myRange.End - myRange.Start 
myRange.Collapse Direction:=wdCollapseStart 
myRange.MoveStartUntil Cset:="$", Count:=leng