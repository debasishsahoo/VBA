Set Range1 = ActiveDocument.Paragraphs(1).Range 
With Range1 
 .Collapse Direction:=wdCollapseStart 
 .Move Unit:=wdParagraph, Count:=3 
 .Select 
End With