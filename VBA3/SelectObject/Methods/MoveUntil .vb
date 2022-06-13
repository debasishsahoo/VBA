x = Selection.MoveUntil(Cset:=Chr$(13), Count:=wdForward) 
MsgBox x-1 & " character positions were moved"