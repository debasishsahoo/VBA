With Selection 
 .StartIsActive = False 
 .MoveRight Unit:=wdWord, Count:=2, Extend:=wdExtend 
End With

With Selection 
 If (.Flags And wdSelStartActive) = wdSelStartActive Then _ 
 .Flags = wdSelReplace 
 .MoveRight Unit:=wdWord, Count:=2, Extend:=wdExtend 
End With

With Selection 
 .MoveEnd Unit:=wdWord, Count:=2 
End With