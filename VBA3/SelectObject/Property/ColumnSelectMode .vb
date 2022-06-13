With Selection 
 .Collapse Direction:=wdCollapseStart 
 .ColumnSelectMode = True 
 .MoveRight Unit:=wdWord, Count:=2, Extend:=wdExtend 
 .MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend 
 .Copy 
 .ColumnSelectMode = False 
End With