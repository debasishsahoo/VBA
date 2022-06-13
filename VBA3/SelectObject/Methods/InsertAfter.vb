ActiveDocument.Range( _ 
 ActiveDocument.Paragraphs(1).Range.Start, _ 
 ActiveDocument.Paragraphs(1).Range.End - 1) _ 
 .Select 
 
Selection.InsertAfter _ 
 " This is now the last sentence in paragraph one."




 With Selection 
 .InsertAfter "appended text" 
 .Collapse Direction:=wdCollapseEnd 
End With