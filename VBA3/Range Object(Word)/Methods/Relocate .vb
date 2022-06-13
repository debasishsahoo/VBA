theStart = ActiveDocument.Paragraphs(3).Range.Start 
theEnd = ActiveDocument.Paragraphs(5).Range.End 
Set myRange = ActiveDocument.Range(Start:=theStart, End:=theEnd) 
ActiveDocument.ActiveWindow.View.Type = wdOutlineView 
myRange.Relocate Direction:=wdRelocateDown


ActiveDocument.ActiveWindow.View.Type = wdOutlineView 
Selection.Paragraphs(1).Range.Relocate Direction:=wdRelocateUp