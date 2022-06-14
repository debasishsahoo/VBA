Sub ReturnTableText() 
 Dim tblOne As Table 
 Dim celTable As Cell 
 Dim rngTable As Range 
 
 Set tblOne = ActiveDocument.Tables(1) 
 For Each celTable In tblOne.Rows(1).Cells 
  Set rngTable = ActiveDocument.Range(Start:=celTable.Range.Start, _ 
  End:=celTable.Range.End - 1) 
  MsgBox rngTable.Text 
 Next celTable 
End Sub

Sub ReturnCellText() 
 Dim tblOne As Table 
 Dim celTable As Cell 
 Dim rngTable As Range 
 
 Set tblOne = ActiveDocument.Tables(1) 
 For Each celTable In tblOne.Rows(1).Cells 
  Set rngTable = celTable.Range 
  rngTable.MoveEnd Unit:=wdCharacter, Count:=-1 
  MsgBox rngTable.Text 
 Next celTable 
End Sub