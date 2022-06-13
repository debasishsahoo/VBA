Set aDoc = ActiveDocument 
Set myRange = aDoc.Range(Start:=aDoc.Paragraphs(1).Range.Start, _ 
 End:=aDoc.Paragraphs(3).Range.End) 
myRange.ConvertToTable Separator:=wdSeparateByParagraphs