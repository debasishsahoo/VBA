If ActiveDocument.Words.Count >= 2 Then 
 Set myRange = ActiveDocument.Words(2) 
 With myRange 
 .MoveStart Unit:=wdWord, Count:=-1 
 .Select 
 End With 
End If