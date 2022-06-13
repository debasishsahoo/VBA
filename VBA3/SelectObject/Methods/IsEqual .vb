If Selection.IsEqual(ActiveDocument _ 
 .Paragraphs(2).Range) = False Then 
 ActiveDocument.Paragraphs(2).Range.Select 
End If