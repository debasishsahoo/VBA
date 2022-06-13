With Selection.Range 
 .MoveEndUntil Cset:="a", Count:=wdForward 
 .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend 
End With