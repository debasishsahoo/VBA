ActiveDocument.Paragraphs(1).Range.Select 
Selection.CopyFormat 
ActiveDocument.Paragraphs(2).Range.Select 
Selection.PasteFormat


With Selection 
 .Collapse Direction:=wdCollapseStart 
 .CopyFormat 
 .Next(Unit:=wdWord, Count:=1).Select 
 .PasteFormat 
End With