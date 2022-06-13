With Selection 
 Set MyRange = .GoTo(wdGoToField, wdGoToPrevious) 
 .MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend 
 If Selection.Fields.Count = 1 Then Selection.Fields(1).Update 
End With

If Selection.MoveRight = 1 Then MsgBox "Move was successful"

If Selection.Information(wdWithInTable) = True Then 
 Selection.MoveRight Unit:=wdCell, Count:=1, Extend:=wdMove 
End If