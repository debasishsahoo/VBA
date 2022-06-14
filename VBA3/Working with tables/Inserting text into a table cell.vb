Sub InsertTextInCell() 
 If ActiveDocument.Tables.Count >= 1 Then 
  With ActiveDocument.Tables(1).Cell(Row:=1, Column:=1).Range 
   .Delete 
   .InsertAfter Text:="Cell 1,1" 
  End With 
 End If 
End Sub