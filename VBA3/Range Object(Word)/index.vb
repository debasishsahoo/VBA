Set myRange = ActiveDocument.Range(Start:=0, End:=10)

Set aRange = ActiveDocument.Paragraphs(1).Range

Set aRange = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(2).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End)