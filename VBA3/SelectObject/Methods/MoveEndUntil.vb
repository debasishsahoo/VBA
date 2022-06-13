With Selection 
 .MoveEndUntil Cset:="a", Count:=wdForward 
 .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend 
End With

char = Selection.MoveEndUntil(Cset:=vbTab, Count:=100) 
If char = 0 Then StatusBar = "Selection not moved"