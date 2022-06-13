Set myRange = ActiveDocument.Content 
myRange.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Tables.Add Range:=myRange, NumRows:=2, NumColumns:=2