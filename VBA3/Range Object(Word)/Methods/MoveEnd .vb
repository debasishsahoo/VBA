If ActiveDocument.Words.Count >= 3 Then 
 Set myRange = ActiveDocument.Words(2) 
 With myRange 
 .MoveEnd Unit:=wdWord, Count:=1 
 .Select 
 End With 
End If