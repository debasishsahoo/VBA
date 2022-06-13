Set myRange = ActiveDocument.Paragraphs(2).Range 
myRange.Select 
Selection.SelectCurrentTabs


With Selection 
 .SelectCurrentTabs 
 pos = .Paragraphs.TabStops(1).Position 
 .Collapse Direction:=wdCollapseEnd 
 .SelectCurrentTabs 
 .Paragraphs.TabStops.Add Position:=pos 
End With

