Sub ReturnCellContentsToArray() 
 Dim intCells As Integer 
 Dim celTable As Cell 
 Dim strCells() As String 
 Dim intCount As Integer 
 Dim rngText As Range 
 
 If ActiveDocument.Tables.Count >= 1 Then 
  With ActiveDocument.Tables(1).Range 
   intCells = .Cells.Count 
   ReDim strCells(intCells) 
   intCount = 1 
   For Each celTable In .Cells 
    Set rngText = celTable.Range 
    rngText.MoveEnd Unit:=wdCharacter, Count:=-1 
    strCells(intCount) = rngText 
    intCount = intCount + 1 
   Next celTable 
  End With 
 End If 
End Sub