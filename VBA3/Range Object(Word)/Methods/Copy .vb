ActiveDocument.Paragraphs(1).Range.Copy 
Set myRange = ActiveDocument.Range _ 
 (Start:=ActiveDocument.Content.End - 1, _ 
 End:=ActiveDocument.Content.End - 1) 
myRange.Paste