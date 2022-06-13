Set myRange = ActiveDocument.Paragraphs(2).Range 
With myRange 
 .Collapse Direction:=wdCollapseEnd 
 .InsertBreak Type:=wdPageBreak 
End With