With Selection 
 .Paragraphs(1).Range.Select 
 .CopyFormat 
 .Paragraphs(1).Next.Range.Select 
 .PasteFormat 
End With

With Selection 
 .Collapse Direction:=wdCollapseStart 
 .CopyFormat 
 .Next(Unit:=wdWord, Count:=1).Select 
 .PasteFormat 
End With
