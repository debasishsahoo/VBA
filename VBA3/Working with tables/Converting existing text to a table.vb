Sub ConvertExistingText() 
 With Documents.Add.Content 
  .InsertBefore "one" & vbTab & "two" & vbTab & "three" & vbCr 
  .ConvertToTable Separator:=Chr(9), NumRows:=1, NumColumns:=3 
 End With 
End Sub